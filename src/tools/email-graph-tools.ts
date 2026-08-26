/**
 * email-graph-tools.ts
 *
 * All email operations implemented directly via Microsoft Graph API.
 * Works with any standard M365 subscription — no Copilot license required.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [email-graph] ${msg}\n`);
}

// ── Graph fetch helper ────────────────────────────────────────────────────────

async function gf(
  token: string,
  path: string,
  options: RequestInit = {}
): Promise<{ ok: boolean; status: number; body: unknown }> {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...((options.headers as Record<string, string>) ?? {}),
    },
  });
  const ct = res.headers.get("content-type") ?? "";
  const body = ct.includes("json") || ct.includes("odata")
    ? await res.json().catch(() => null)
    : await res.text().catch(() => "");
  return { ok: res.ok, status: res.status, body };
}

// ── Recipient helpers ─────────────────────────────────────────────────────────

function toRecipientList(
  val: unknown
): Array<{ emailAddress: { address: string; name?: string } }> {
  if (!val) return [];
  const arr = Array.isArray(val) ? val : String(val).split(/[,;]/).map(s => s.trim()).filter(Boolean);
  return arr
    .map((r) => {
      if (typeof r !== "string") return { emailAddress: { address: String(r) } };
      const m = r.match(/^(.+?)\s*<([^>]+)>$/);
      if (m) return { emailAddress: { name: m[1].trim(), address: m[2].trim() } };
      return { emailAddress: { address: r.trim() } };
    })
    .filter((r) => r.emailAddress.address.includes("@"));
}

// ── Message formatter ─────────────────────────────────────────────────────────

function fmtMsg(m: Record<string, unknown>): string {
  const from = (m.from as { emailAddress?: { name?: string; address?: string } } | undefined)
    ?.emailAddress;
  const to = ((m.toRecipients as Array<{ emailAddress?: { name?: string; address?: string } }>) ?? [])
    .map((r) => r?.emailAddress?.name ?? r?.emailAddress?.address ?? "")
    .filter(Boolean)
    .join(", ");
  const bodyContent = ((m.body as { content?: string } | undefined)?.content ?? "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, 600);
  return [
    `ID: ${m.id}`,
    `Subject: ${m.subject ?? "(no subject)"}`,
    `From: ${from?.name ? `${from.name} <${from.address ?? ""}>` : (from?.address ?? "")}`,
    `To: ${to}`,
    `Date: ${m.receivedDateTime ?? m.sentDateTime ?? ""}`,
    `Read: ${m.isRead ? "Yes" : "No"}`,
    `Has Attachments: ${m.hasAttachments ? "Yes" : "No"}`,
    bodyContent ? `\nBody:\n${bodyContent}` : (m.bodyPreview ? `\nPreview: ${m.bodyPreview}` : ""),
  ]
    .filter(Boolean)
    .join("\n");
}

// ── Tool names set (used by proxy server for routing) ─────────────────────────

export const EMAIL_GRAPH_TOOL_NAMES = new Set([
  "SearchMessages",
  "UploadLargeAttachment",
  "SearchMessagesQueryParameters",
  "GetMessage",
  "CreateDraftMessage",
  "SendDraftMessage",
  "SendEmailWithAttachments",
  "UpdateDraft",
  "DeleteMessage",
  "UpdateMessage",
  "FlagEmail",
  "ReplyToMessage",
  "ReplyAllToMessage",
  "ReplyWithFullThread",
  "ReplyAllWithFullThread",
  "ForwardMessage",
  "ForwardMessageWithFullThread",
  "GetAttachments",
  "DownloadAttachment",
  "UploadAttachment",
  "AddDraftAttachments",
  "DeleteAttachment",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const EMAIL_GRAPH_TOOLS: Tool[] = [
  {
    name: "SearchMessages",
    description:
      "Search for emails in the mailbox. Returns matching messages with ID, subject, sender, date, and body.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Full-text search (subject, body, sender)" },
        folder: {
          type: "string",
          description: "Folder to search: inbox (default), sentitems, drafts, deleteditems",
        },
        top: { type: "number", description: "Max results to return (default: 10, max: 50)" },
        filter: { type: "string", description: "OData $filter expression" },
        orderby: {
          type: "string",
          description: "Sort order (default: receivedDateTime desc)",
        },
      },
    },
  },
  {
    name: "SearchMessagesQueryParameters",
    description: "Search mailbox messages by subject, sender, date range, read status, or attachments.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Subject contains this text" },
        from: { type: "string", description: "Filter by sender email address" },
        to: { type: "string", description: "Filter by recipient email address" },
        hasAttachment: { type: "boolean", description: "Only messages with attachments" },
        isRead: { type: "boolean", description: "Filter by read/unread status" },
        receivedAfter: { type: "string", description: "ISO 8601 — received after this datetime" },
        receivedBefore: { type: "string", description: "ISO 8601 — received before this datetime" },
        folder: {
          type: "string",
          description: "Folder: inbox (default), sentitems, drafts, deleteditems",
        },
        top: { type: "number", description: "Max results (default: 10)" },
      },
    },
  },
  {
    name: "GetMessage",
    description: "Get the full content of a specific email message by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "The message ID" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "CreateDraftMessage",
    description: "Create a new draft email without sending it. Returns the draft ID.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body content" },
        contentType: { type: "string", description: "Body type: 'HTML' or 'text' (default: text)" },
        to: {
          description: "To recipients — single address or array",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        cc: {
          description: "CC recipients",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        bcc: {
          description: "BCC recipients",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
      },
    },
  },
  {
    name: "UpdateDraft",
    description: "Update an existing draft message — change recipients, subject, or body.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the draft to update" },
        subject: { type: "string" },
        body: { type: "string" },
        contentType: { type: "string", description: "'HTML' or 'text'" },
        to: {
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        cc: {
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        bcc: {
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
      },
      required: ["messageId"],
    },
  },
  {
    name: "DeleteMessage",
    description: "Delete an email message from the mailbox.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message to delete" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "UpdateMessage",
    description: "Update mutable message properties such as read status or categories.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message" },
        isRead: { type: "boolean", description: "Mark as read (true) or unread (false)" },
        categories: { type: "array", items: { type: "string" }, description: "Categories to set" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "FlagEmail",
    description: "Set the flag status on an email: flagged, complete, or notFlagged.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message" },
        flagStatus: {
          type: "string",
          description: "Flag status: 'flagged', 'complete', or 'notFlagged'",
        },
      },
      required: ["messageId", "flagStatus"],
    },
  },
  {
    name: "ReplyToMessage",
    description: "Reply to an email message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message to reply to" },
        comment: { type: "string", description: "Reply body text" },
        body: { type: "string", description: "Alternative to 'comment'" },
        contentType: { type: "string", description: "'HTML' or 'text'" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "ReplyAllToMessage",
    description: "Reply to all recipients of an email message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message" },
        comment: { type: "string", description: "Reply body text" },
        body: { type: "string" },
        contentType: { type: "string" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "ReplyWithFullThread",
    description: "Reply to a message, preserving the full quoted thread.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        comment: { type: "string" },
        contentType: { type: "string" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "ReplyAllWithFullThread",
    description: "Reply-all to a message, preserving the full quoted thread.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        comment: { type: "string" },
        contentType: { type: "string" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "ForwardMessage",
    description: "Forward an email to other recipients.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message to forward" },
        to: {
          description: "Recipients to forward to",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        toRecipients: {
          type: "array",
          items: { type: "string" },
          description: "Alternative to 'to'",
        },
        comment: { type: "string", description: "Optional forwarding note" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "ForwardMessageWithFullThread",
    description: "Forward a message including the full email thread.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        to: {
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        comment: { type: "string" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "GetAttachments",
    description: "List all attachments on a message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "DownloadAttachment",
    description: "Download a specific attachment from a message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        attachmentId: { type: "string" },
      },
      required: ["messageId", "attachmentId"],
    },
  },
  {
    name: "UploadAttachment",
    description: "Upload a file attachment to a message or draft (under 3 MB).",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        name: { type: "string", description: "File name" },
        contentType: { type: "string", description: "MIME type (e.g. application/pdf)" },
        contentBytes: { type: "string", description: "Base64-encoded file content" },
      },
      required: ["messageId", "name", "contentBytes"],
    },
  },
  {
    name: "AddDraftAttachments",
    description: "Add one or more file attachments to an existing draft message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the draft" },
        attachments: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              contentType: { type: "string" },
              contentBytes: { type: "string", description: "Base64-encoded content" },
            },
            required: ["name", "contentBytes"],
          },
        },
      },
      required: ["messageId", "attachments"],
    },
  },
  {
    name: "DeleteAttachment",
    description: "Delete an attachment from a message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        attachmentId: { type: "string" },
      },
      required: ["messageId", "attachmentId"],
    },
  },
  {
    name: "SendDraftMessage",
    description: "Send an existing draft email message.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the draft message to send" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "SendEmailWithAttachments",
    description:
      "Create and immediately send an email, optionally attaching files straight from the user's " +
      "OneDrive or a SharePoint site. To attach a file, pass its path or item ID in `attachments` — " +
      "the file's bytes are fetched and attached for you, so there is no need to download or " +
      "base64-encode anything first. The user's signature is appended automatically.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body content" },
        contentType: { type: "string", description: "Body type: 'HTML' or 'text' (default: text)" },
        to: {
          description: "To recipients",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        cc: {
          description: "CC recipients",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        bcc: {
          description: "BCC recipients",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        attachments: {
          type: "array",
          description:
            "Optional. Files to attach, each identified by its OneDrive path (e.g. " +
            "'/Reports/Q3.pptx' or just 'Q3.pptx' for a file in the root) or by the item ID " +
            "returned from list_onedrive_folder. Use siteId on an entry to pull the file from a " +
            "SharePoint site instead of the personal OneDrive.",
          items: {
            type: "object",
            properties: {
              ref: { type: "string", description: "File path relative to the drive root, or the file's item ID." },
              siteId: { type: "string", description: "Optional. SharePoint site ID, when the file lives in a site rather than the user's OneDrive." },
            },
            required: ["ref"],
          },
        },
      },
    },
  },
  {
    name: "UploadLargeAttachment",
    description: "Upload a large file attachment (3 MB – 150 MB) to a message or draft using a resumable upload session.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of the message or draft" },
        name: { type: "string", description: "Attachment file name" },
        contentType: { type: "string", description: "MIME type (e.g. application/pdf)" },
        contentBytes: { type: "string", description: "Base64-encoded file content" },
      },
      required: ["messageId", "name", "contentBytes"],
    },
  },
];

// ── Handler class ─────────────────────────────────────────────────────────────

export class EmailGraphToolHandler {
  private token: string;

  constructor(token: string) {
    this.token = token;
  }

  async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    try {
      const text = await this.dispatch(name, args);
      const isError = /^(Error|Unknown)/.test(text);
      return { content: [{ type: "text", text }], ...(isError ? { isError: true } : {}) };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      log(`${name} error: ${msg}`);
      return { content: [{ type: "text", text: `Error: ${msg}` }], isError: true };
    }
  }

  private async dispatch(name: string, args: Record<string, unknown>): Promise<string> {
    switch (name) {
      case "SearchMessages":               return this.search(args);
      case "SearchMessagesQueryParameters": return this.searchQuery(args);
      case "GetMessage":                   return this.get(args);
      case "CreateDraftMessage":           return this.createDraft(args);
      case "SendDraftMessage":             return this.sendDraft(args);
      case "SendEmailWithAttachments":     return this.sendEmail(args);
      case "UpdateDraft":                  return this.updateDraft(args);
      case "DeleteMessage":                return this.del(args);
      case "UpdateMessage":                return this.update(args);
      case "FlagEmail":                    return this.flag(args);
      case "ReplyToMessage":
      case "ReplyWithFullThread":          return this.reply(args, false);
      case "ReplyAllToMessage":
      case "ReplyAllWithFullThread":       return this.replyAll(args, false);
      case "ForwardMessage":
      case "ForwardMessageWithFullThread": return this.forward(args);
      case "GetAttachments":               return this.getAttachments(args);
      case "DownloadAttachment":           return this.downloadAttachment(args);
      case "UploadAttachment":             return this.uploadAttachment(args);
      case "UploadLargeAttachment":        return this.uploadLargeAttachment(args);
      case "AddDraftAttachments":          return this.addDraftAttachments(args);
      case "DeleteAttachment":             return this.deleteAttachment(args);
      default: return `Unknown email tool: ${name}`;
    }
  }

  // ── Search ──────────────────────────────────────────────────────────────────

  private async search(args: Record<string, unknown>): Promise<string> {
    const folder = String(args.folder ?? "inbox");
    const top = Math.min(Number(args.top ?? 10), 50);
    const folderPath = this.folderPath(folder);
    const params = new URLSearchParams({
      $top: String(top),
      $orderby: String(args.orderby ?? "receivedDateTime desc"),
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,body",
    });
    if (args.query) params.set("$search", `"${args.query}"`);
    if (args.filter) params.set("$filter", String(args.filter));

    const res = await gf(this.token, `/me${folderPath}?${params}`);
    if (!res.ok) return `Error searching messages (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body as { value?: unknown[] };
    const msgs = data.value ?? [];
    if (msgs.length === 0) return "No messages found.";
    return msgs.map((m) => fmtMsg(m as Record<string, unknown>)).join("\n\n---\n\n");
  }

  private async searchQuery(args: Record<string, unknown>): Promise<string> {
    const filters: string[] = [];
    if (args.from) filters.push(`from/emailAddress/address eq '${args.from}'`);
    if (args.to) filters.push(`toRecipients/any(r: r/emailAddress/address eq '${args.to}')`);
    if (typeof args.isRead === "boolean") filters.push(`isRead eq ${args.isRead}`);
    if (typeof args.hasAttachment === "boolean") filters.push(`hasAttachments eq ${args.hasAttachment}`);
    if (args.receivedAfter) filters.push(`receivedDateTime ge ${args.receivedAfter}`);
    if (args.receivedBefore) filters.push(`receivedDateTime le ${args.receivedBefore}`);

    const top = Math.min(Number(args.top ?? 10), 50);
    const folderPath = this.folderPath(String(args.folder ?? "inbox"));
    const params = new URLSearchParams({
      $top: String(top),
      $orderby: "receivedDateTime desc",
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,body",
    });
    if (filters.length > 0) params.set("$filter", filters.join(" and "));
    if (args.subject) params.set("$search", `"${args.subject}"`);

    const res = await gf(this.token, `/me${folderPath}?${params}`);
    if (!res.ok) return `Error searching messages (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body as { value?: unknown[] };
    const msgs = data.value ?? [];
    if (msgs.length === 0) return "No messages found matching your criteria.";
    return msgs.map((m) => fmtMsg(m as Record<string, unknown>)).join("\n\n---\n\n");
  }

  private folderPath(folder: string): string {
    const map: Record<string, string> = {
      inbox: "/mailFolders/inbox/messages",
      sentitems: "/mailFolders/sentitems/messages",
      drafts: "/mailFolders/drafts/messages",
      deleteditems: "/mailFolders/deleteditems/messages",
    };
    return map[folder.toLowerCase()] ?? "/messages";
  }

  // ── Get ─────────────────────────────────────────────────────────────────────

  private async get(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const res = await gf(this.token, `/me/messages/${id}?$expand=attachments`);
    if (!res.ok) return `Error getting message (${res.status}): ${JSON.stringify(res.body)}`;
    return fmtMsg(res.body as Record<string, unknown>);
  }

  // ── Draft ───────────────────────────────────────────────────────────────────

  private async createDraft(args: Record<string, unknown>): Promise<string> {
    const ct = String(args.contentType ?? args.bodyType ?? "text");
    const payload = {
      subject: String(args.subject ?? ""),
      body: {
        contentType: ct.toLowerCase() === "html" ? "HTML" : "Text",
        content: String(args.body ?? args.emailBody ?? ""),
      },
      toRecipients: toRecipientList(args.to ?? args.toRecipients),
      ccRecipients: toRecipientList(args.cc ?? args.ccRecipients),
      bccRecipients: toRecipientList(args.bcc ?? args.bccRecipients),
    };
    const res = await gf(this.token, "/me/messages", { method: "POST", body: JSON.stringify(payload) });
    if (!res.ok) return `Error creating draft (${res.status}): ${JSON.stringify(res.body)}`;
    const d = res.body as { id: string; subject?: string };
    return `Draft created.\nID: ${d.id}\nSubject: ${d.subject ?? ""}`;
  }

  private async updateDraft(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const patch: Record<string, unknown> = {};
    if (args.subject !== undefined) patch.subject = args.subject;
    if (args.body !== undefined || args.emailBody !== undefined) {
      const ct = String(args.contentType ?? "text");
      patch.body = {
        contentType: ct.toLowerCase() === "html" ? "HTML" : "Text",
        content: String(args.body ?? args.emailBody ?? ""),
      };
    }
    if (args.to !== undefined) patch.toRecipients = toRecipientList(args.to);
    if (args.cc !== undefined) patch.ccRecipients = toRecipientList(args.cc);
    if (args.bcc !== undefined) patch.bccRecipients = toRecipientList(args.bcc);

    const res = await gf(this.token, `/me/messages/${id}`, { method: "PATCH", body: JSON.stringify(patch) });
    if (!res.ok) return `Error updating draft (${res.status}): ${JSON.stringify(res.body)}`;
    return "Draft updated.";
  }

  // ── Delete / Update / Flag ───────────────────────────────────────────────────

  private async del(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const res = await gf(this.token, `/me/messages/${id}`, { method: "DELETE" });
    if (!res.ok) return `Error deleting message (${res.status}): ${JSON.stringify(res.body)}`;
    return "Message deleted.";
  }

  private async update(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const patch: Record<string, unknown> = {};
    if (args.isRead !== undefined) patch.isRead = args.isRead;
    if (args.categories !== undefined) patch.categories = args.categories;
    const res = await gf(this.token, `/me/messages/${id}`, { method: "PATCH", body: JSON.stringify(patch) });
    if (!res.ok) return `Error updating message (${res.status}): ${JSON.stringify(res.body)}`;
    return "Message updated.";
  }

  private async flag(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const status = String(args.flagStatus ?? args.status ?? "flagged");
    const res = await gf(this.token, `/me/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ flag: { flagStatus: status } }),
    });
    if (!res.ok) return `Error flagging email (${res.status}): ${JSON.stringify(res.body)}`;
    return `Email flagged as "${status}".`;
  }

  // ── Reply / Reply-All / Forward ─────────────────────────────────────────────

  private async reply(args: Record<string, unknown>, _full: boolean): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const comment = String(args.comment ?? args.body ?? "");
    const res = await gf(this.token, `/me/messages/${id}/reply`, {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
    if (!res.ok) return `Error sending reply (${res.status}): ${JSON.stringify(res.body)}`;
    return "Reply sent successfully.";
  }

  private async replyAll(args: Record<string, unknown>, _full: boolean): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const comment = String(args.comment ?? args.body ?? "");
    const res = await gf(this.token, `/me/messages/${id}/replyAll`, {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
    if (!res.ok) return `Error sending reply-all (${res.status}): ${JSON.stringify(res.body)}`;
    return "Reply-all sent successfully.";
  }

  private async forward(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const toRecipients = toRecipientList(args.to ?? args.toRecipients);
    const res = await gf(this.token, `/me/messages/${id}/forward`, {
      method: "POST",
      body: JSON.stringify({ comment: String(args.comment ?? ""), toRecipients }),
    });
    if (!res.ok) return `Error forwarding message (${res.status}): ${JSON.stringify(res.body)}`;
    return "Message forwarded.";
  }

  // ── Attachments ─────────────────────────────────────────────────────────────

  private async getAttachments(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.id ?? "");
    const res = await gf(this.token, `/me/messages/${id}/attachments`);
    if (!res.ok) return `Error getting attachments (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body as { value?: Array<{ id: string; name: string; size: number; contentType: string }> };
    const atts = data.value ?? [];
    if (atts.length === 0) return "No attachments.";
    return atts.map((a) => `ID: ${a.id}\nName: ${a.name}\nSize: ${a.size} bytes\nType: ${a.contentType}`).join("\n\n");
  }

  private async downloadAttachment(args: Record<string, unknown>): Promise<string> {
    const msgId = String(args.messageId ?? args.id ?? "");
    const attId = String(args.attachmentId ?? "");
    const res = await gf(this.token, `/me/messages/${msgId}/attachments/${attId}`);
    if (!res.ok) return `Error downloading attachment (${res.status}): ${JSON.stringify(res.body)}`;
    const a = res.body as { name?: string; contentType?: string; contentBytes?: string; size?: number };
    return `Name: ${a.name}\nType: ${a.contentType}\nSize: ${a.size} bytes\nContent (base64):\n${a.contentBytes ?? "(empty)"}`;
  }

  private async uploadAttachment(args: Record<string, unknown>): Promise<string> {
    const msgId = String(args.messageId ?? args.id ?? "");
    const res = await gf(this.token, `/me/messages/${msgId}/attachments`, {
      method: "POST",
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: args.name,
        contentType: args.contentType ?? "application/octet-stream",
        contentBytes: args.contentBytes,
      }),
    });
    if (!res.ok) return `Error uploading attachment (${res.status}): ${JSON.stringify(res.body)}`;
    return "Attachment uploaded.";
  }

  private async addDraftAttachments(args: Record<string, unknown>): Promise<string> {
    const msgId = String(args.messageId ?? args.id ?? "");
    const atts = args.attachments as Array<{ name: string; contentType?: string; contentBytes: string }>;
    for (const att of atts) {
      const res = await gf(this.token, `/me/messages/${msgId}/attachments`, {
        method: "POST",
        body: JSON.stringify({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: att.name,
          contentType: att.contentType ?? "application/octet-stream",
          contentBytes: att.contentBytes,
        }),
      });
      if (!res.ok) return `Error adding attachment "${att.name}" (${res.status}): ${JSON.stringify(res.body)}`;
    }
    return `${atts.length} attachment(s) added.`;
  }

  private async uploadLargeAttachment(args: Record<string, unknown>): Promise<string> {
    const msgId = String(args.messageId ?? args.id ?? "");
    const name = String(args.name ?? "attachment");
    const contentType = String(args.contentType ?? "application/octet-stream");
    const b64 = String(args.contentBytes ?? "");
    const buffer = Buffer.from(b64, "base64");
    const totalSize = buffer.length;

    // 1️⃣ Create upload session
    const sessionRes = await gf(this.token, `/me/messages/${msgId}/attachments/createUploadSession`, {
      method: "POST",
      body: JSON.stringify({
        AttachmentItem: { attachmentType: "file", name, size: totalSize, contentType },
      }),
    });
    if (!sessionRes.ok)
      return `Error creating upload session (${sessionRes.status}): ${JSON.stringify(sessionRes.body)}`;
    const { uploadUrl } = sessionRes.body as { uploadUrl?: string };
    if (!uploadUrl) return "Error: no upload URL returned by Graph";

    // 2️⃣ Upload in 4 MB chunks
    const CHUNK = 4 * 1024 * 1024;
    let offset = 0;
    while (offset < totalSize) {
      const end = Math.min(offset + CHUNK, totalSize);
      const chunk = buffer.slice(offset, end);
      const chunkRes = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Type": "application/octet-stream",
          "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
          "Content-Length": String(chunk.length),
        },
        body: chunk,
      });
      if (!chunkRes.ok && chunkRes.status !== 200 && chunkRes.status !== 201) {
        const err = await chunkRes.text();
        return `Error uploading chunk at offset ${offset}: ${err}`;
      }
      offset = end;
    }
    return `Large attachment "${name}" (${Math.round(totalSize / 1024)} KB) uploaded successfully.`;
  }

  private async deleteAttachment(args: Record<string, unknown>): Promise<string> {
    const msgId = String(args.messageId ?? args.id ?? "");
    const attId = String(args.attachmentId ?? "");
    const res = await gf(this.token, `/me/messages/${msgId}/attachments/${attId}`, { method: "DELETE" });
    if (!res.ok) return `Error deleting attachment (${res.status}): ${JSON.stringify(res.body)}`;
    return "Attachment deleted.";
  }

  // ── Send ────────────────────────────────────────────────────────────────────

  async sendDraft(args: Record<string, unknown>): Promise<string> {
    const id = String(args.messageId ?? args.draftId ?? args.id ?? "");
    if (!id) return "Error: messageId is required.";
    const res = await gf(this.token, `/me/messages/${id}/send`, { method: "POST", body: "" });
    if (!res.ok) return `Error sending draft (${res.status}): ${JSON.stringify(res.body)}`;
    return "Email sent successfully.";
  }

  async sendEmail(args: Record<string, unknown>): Promise<string> {
    const ct = String(args.contentType ?? args.bodyType ?? "text");
    const message = {
      subject: String(args.subject ?? ""),
      body: {
        contentType: ct.toLowerCase() === "html" ? "HTML" : "Text",
        content: String(args.body ?? args.emailBody ?? ""),
      },
      toRecipients:  toRecipientList(args.to  ?? args.toRecipients),
      ccRecipients:  toRecipientList(args.cc  ?? args.ccRecipients),
      bccRecipients: toRecipientList(args.bcc ?? args.bccRecipients),
    };
    const res = await gf(this.token, "/me/sendMail", { method: "POST", body: JSON.stringify({ message }) });
    if (!res.ok) return `Error sending email (${res.status}): ${JSON.stringify(res.body)}`;
    return `Email sent successfully to ${(message.toRecipients).map(r => r.emailAddress.address).join(", ")}.`;
  }
}
