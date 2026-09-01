import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { fmtDate } from "./format";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
}

// Graph stores whatever Content-Type we send, and Office/Outlook rely on it to
// decide how to open a file — an .xlsx saved as text/plain downloads as a
// broken blob. Mapping by extension keeps callers from having to know MIME
// strings while still letting them override explicitly.
const MIME_TYPES: Record<string, string> = {
  pdf:  "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  doc:  "application/msword",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xls:  "application/vnd.ms-excel",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ppt:  "application/vnd.ms-powerpoint",
  png:  "image/png",
  jpg:  "image/jpeg",
  jpeg: "image/jpeg",
  gif:  "image/gif",
  csv:  "text/csv",
  txt:  "text/plain",
  json: "application/json",
  zip:  "application/zip",
};

function guessMimeType(fileName: string): string {
  const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
  return MIME_TYPES[ext] ?? "application/octet-stream";
}

/**
 * Builds the Graph path for writing a file by name into a folder.
 *
 * Every document-creating tool (Word, Excel, PowerPoint) needs exactly this,
 * and each had its own copy hardcoded to `/me/drive` — which is why none of
 * them could create anything in SharePoint. Keeping the OneDrive/SharePoint
 * decision in one place means adding a format later cannot silently
 * reintroduce that gap.
 *
 * @param folderPath Folder relative to the drive root ('/Reports'); root if omitted.
 * @param siteId     SharePoint site; omit for the user's own OneDrive.
 * @param driveId    A specific document library; otherwise the site's default.
 */
export function buildUploadPath(
  fileName: string,
  folderPath?: string,
  siteId?: string,
  driveId?: string
): string {
  const base = driveId
    ? `/drives/${driveId}`
    : siteId
      ? `/sites/${siteId}/drive`
      : "/me/drive";

  const folder = (folderPath ?? "").replace(/^\/+|\/+$/g, "");
  const encodedName = encodeURIComponent(fileName);
  return folder
    ? `${base}/root:/${encodeURI(folder)}/${encodedName}:/content`
    : `${base}/root:/${encodedName}:/content`;
}

// ── Tool Definitions ──────────────────────────────────────────────────────────

export const SHAREPOINT_TOOLS: Tool[] = [
  {
    name: "create_sharepoint_folder",
    description: "Creates a new folder in a SharePoint site. First use getSiteByPath or searchSitesByName to get the siteId. The driveId can be found by listing the site drives.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "SharePoint site ID from getSiteByPath or searchSitesByName.",
        },
        parentItemId: {
          type: "string",
          description: "Item ID of the parent folder. Use 'root' for the default document library root. Get item IDs from list_sharepoint_folder.",
        },
        folderName: {
          type: "string",
          description: "Name of the new folder to create.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID for a specific document library. If omitted, uses the site's default drive.",
        },
      },
      required: ["siteId", "parentItemId", "folderName"],
    },
  },
  {
    name: "list_sharepoint_folder",
    description: "Lists the contents of a SharePoint folder by item ID. Use 'root' to list the default document library root.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "SharePoint site ID.",
        },
        itemId: {
          type: "string",
          description: "Item ID of the folder to list. Use 'root' for the document library root.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID for a specific document library.",
        },
      },
      required: ["siteId", "itemId"],
    },
  },
  {
    name: "upload_file_to_sharepoint",
    description:
      "Uploads a file to a SharePoint folder. Supply 'content' for plain text, or 'contentBase64' " +
      "for any other format (PDF, Word, Excel, PowerPoint, images). Uploading to a name that " +
      "already exists replaces that file's contents.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "SharePoint site ID.",
        },
        parentItemId: {
          type: "string",
          description: "Item ID of the destination folder. Use 'root' for document library root.",
        },
        fileName: {
          type: "string",
          description: "File name including extension e.g. 'report.pdf'.",
        },
        content: {
          type: "string",
          description: "Text content, for plain-text files. Ignored when contentBase64 is supplied.",
        },
        contentBase64: {
          type: "string",
          description: "Base64-encoded file bytes, for any non-text format. Takes precedence over 'content'.",
        },
        contentType: {
          type: "string",
          description: "Optional MIME type. Inferred from the file extension when omitted.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID for a specific document library.",
        },
      },
      required: ["siteId", "parentItemId", "fileName"],
    },
  },
  {
    name: "move_sharepoint_file",
    description: "Moves a file or folder to a different parent folder in SharePoint.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "SharePoint site ID.",
        },
        itemId: {
          type: "string",
          description: "Item ID of the file or folder to move.",
        },
        destinationItemId: {
          type: "string",
          description: "Item ID of the destination folder.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID.",
        },
      },
      required: ["siteId", "itemId", "destinationItemId"],
    },
  },
  {
    name: "delete_sharepoint_file",
    description: "Permanently deletes a file or folder from SharePoint by item ID.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "SharePoint site ID.",
        },
        itemId: {
          type: "string",
          description: "Item ID of the file or folder to delete.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID.",
        },
      },
      required: ["siteId", "itemId"],
    },
  },
  {
    name: "create_onedrive_folder",
    description: "Creates a new folder in the user's OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        parentItemId: {
          type: "string",
          description: "Item ID of the parent folder. Use 'root' for OneDrive root.",
        },
        folderName: {
          type: "string",
          description: "Name of the folder to create.",
        },
      },
      required: ["parentItemId", "folderName"],
    },
  },
  {
    name: "upload_file_to_onedrive",
    description:
      "Uploads a file to the user's OneDrive. Supply 'content' for plain text, or 'contentBase64' " +
      "for any other format (PDF, Word, Excel, PowerPoint, images). Uploading to a name that " +
      "already exists replaces that file's contents.",
    inputSchema: {
      type: "object",
      properties: {
        parentItemId: {
          type: "string",
          description: "Item ID of the destination folder. Use 'root' for OneDrive root.",
        },
        fileName: {
          type: "string",
          description: "File name including extension e.g. 'notes.txt' or 'deck.pptx'.",
        },
        content: {
          type: "string",
          description: "Text content, for plain-text files. Ignored when contentBase64 is supplied.",
        },
        contentBase64: {
          type: "string",
          description: "Base64-encoded file bytes, for any non-text format. Takes precedence over 'content'.",
        },
        contentType: {
          type: "string",
          description: "Optional MIME type. Inferred from the file extension when omitted.",
        },
      },
      required: ["parentItemId", "fileName"],
    },
  },
  {
    name: "read_text_file",
    description:
      "Read the contents of a text-based file from the user's OneDrive or a SharePoint site — " +
      ".txt, .csv, .json, .md, .log, .xml, .yaml, .srt, and .vtt (Teams transcripts and subtitles). " +
      "Identify the file by path (e.g. '/Reports/data.csv') or by the item ID from a folder " +
      "listing. For Word, Excel, PowerPoint or PDF files use their own tools instead — this one " +
      "reads raw text and will refuse those formats.",
    inputSchema: {
      type: "object",
      properties: {
        file: {
          type: "string",
          description: "The file: a path relative to the drive root, or its item ID.",
        },
        siteId: {
          type: "string",
          description: "Optional. SharePoint site ID, when the file lives in a site rather than the user's OneDrive.",
        },
        maxChars: {
          type: "number",
          description:
            "Optional cap on how much text to return (default 100000). If the file is longer the " +
            "result is truncated and says so, along with the full size.",
        },
      },
      required: ["file"],
    },
  },
  {
    name: "transfer_drive_item",
    description:
      "Copy or move a file or folder between ANY two locations the user can reach: OneDrive to " +
      "SharePoint, SharePoint to OneDrive, between two different SharePoint sites, or within the " +
      "same drive. This is the only tool that can cross between drives — move_onedrive_file and " +
      "move_sharepoint_file only work inside one drive. Locations are given as paths (e.g. " +
      "'/Reports/Q3.pptx') or item IDs; omit the site ID for the user's personal OneDrive and " +
      "supply it to address a SharePoint site. Set operation to 'move' to remove the original " +
      "after a successful copy — confirm with the user first, as that deletes from the source.",
    inputSchema: {
      type: "object",
      properties: {
        source: {
          type: "string",
          description: "The file or folder to transfer: a path relative to its drive root, or its item ID.",
        },
        sourceSiteId: {
          type: "string",
          description: "Optional. SharePoint site ID holding the source. Omit for the user's OneDrive.",
        },
        destinationFolder: {
          type: "string",
          description: "Destination FOLDER (not file): a path or item ID. Defaults to the destination drive's root.",
        },
        destinationSiteId: {
          type: "string",
          description: "Optional. SharePoint site ID to transfer into. Omit for the user's OneDrive.",
        },
        newName: {
          type: "string",
          description: "Optional. Name for the copy at the destination. Defaults to the original name.",
        },
        operation: {
          type: "string",
          enum: ["copy", "move"],
          description: "'copy' (default) leaves the original in place. 'move' deletes it once the copy succeeds.",
        },
      },
      required: ["source"],
    },
  },
  {
    name: "list_onedrive_folder",
    description:
      "Lists the folders and files in the user's own OneDrive. This is the tool for questions like " +
      "'what is in my OneDrive'. Call it with no arguments to list the OneDrive root. To look inside " +
      "a specific folder, pass either its itemId (returned by a previous listing) or its path — the " +
      "path form avoids having to walk the tree folder by folder. Note this is the personal OneDrive, " +
      "not a SharePoint site: use list_sharepoint_folder for those.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID of the folder to list. Use 'root' or omit for the OneDrive root. Ignored when 'path' is supplied.",
        },
        path: {
          type: "string",
          description: "Optional. Folder path relative to the OneDrive root, e.g. '/Documents/Reports'. Use this when you know the folder name but not its item ID.",
        },
      },
    },
  },
  {
    name: "move_onedrive_file",
    description:
      "Moves and/or renames a file or folder in the user's OneDrive. Supply destinationItemId to move it, " +
      "newName to rename it, or both at once. Get item IDs from list_onedrive_folder.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID of the file or folder to move or rename.",
        },
        destinationItemId: {
          type: "string",
          description: "Optional. Item ID of the destination folder ('root' for the OneDrive root). Omit to rename in place.",
        },
        newName: {
          type: "string",
          description: "Optional. New name for the item, including extension for files. Omit to keep the current name.",
        },
      },
      required: ["itemId"],
    },
  },
  {
    name: "delete_onedrive_file",
    description:
      "Deletes a file or folder from the user's OneDrive by item ID. The item goes to the OneDrive recycle bin " +
      "rather than being destroyed outright, but treat this as destructive and confirm with the user first. " +
      "Deleting a folder also removes everything inside it. Get item IDs from list_onedrive_folder.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "Item ID of the file or folder to delete.",
        },
      },
      required: ["itemId"],
    },
  },
];

// ── Tool Implementations ──────────────────────────────────────────────────────

export class SharePointToolHandler {
  private token: string;

  constructor(token: string) {
    this.token = token;
  }

  private async graphRequest(
    path: string,
    method = "GET",
    body?: unknown
  ): Promise<unknown> {
    const url = `https://graph.microsoft.com/v1.0${path}`;
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "application/json",
      },
      ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
    });

    if (res.status === 204) return { success: true };
    const text = await res.text();
    if (!res.ok) {
      throw new Error(`Graph API ${res.status}: ${text}`);
    }
    return JSON.parse(text);
  }

  // Accepts a Buffer as well as a string so any file type can be written, not
  // just text. Graph's simple PUT upload covers files up to 250 MB, so no
  // resumable session is needed here (unlike mail attachments, which cut off
  // at 3 MB).
  private async graphUpload(
    path: string,
    content: string | Buffer,
    contentType = "text/plain"
  ): Promise<unknown> {
    const url = `https://graph.microsoft.com/v1.0${path}`;
    const res = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": contentType,
      },
      body: typeof content === "string" ? content : new Uint8Array(content),
    });
    const text = await res.text();
    if (!res.ok) {
      throw new Error(`Graph API upload ${res.status}: ${text}`);
    }
    return JSON.parse(text);
  }

  // Resolves the two ways a caller can supply file bytes. Text stays the easy
  // path; contentBase64 is what makes binary formats (pptx, pdf, docx, images)
  // possible at all, and it wins when both are given.
  private resolveUploadBody(
    args: { content?: unknown; contentBase64?: unknown; contentType?: unknown },
    fileName: string
  ): { body: string | Buffer; contentType: string } {
    const b64 = typeof args.contentBase64 === "string" ? args.contentBase64.trim() : "";
    if (b64) {
      return {
        body: Buffer.from(b64, "base64"),
        contentType: String(args.contentType || guessMimeType(fileName)),
      };
    }
    return {
      body: String(args.content ?? ""),
      contentType: String(args.contentType || "text/plain"),
    };
  }

  // ── Resolve the drive base path ───────────────────────────────────────────
  private siteDrivePath(siteId: string, driveId?: string): string {
    return driveId
      ? `/drives/${driveId}`
      : `/sites/${siteId}/drive`;
  }

  // ── SharePoint: Create folder ─────────────────────────────────────────────
  async createSharePointFolder(
    siteId: string,
    parentItemId: string,
    folderName: string,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    const parentPath = parentItemId === "root"
      ? `${base}/root/children`
      : `${base}/items/${parentItemId}/children`;

    await this.graphRequest(parentPath, "POST", {
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });

    log(`Created SharePoint folder: ${folderName} under ${parentItemId}`);
    return `✅ Folder '${folderName}' created successfully.`;
  }

  // ── Shared: render a /children listing ────────────────────────────────────
  // Item IDs are included because every follow-up operation (list a subfolder,
  // move, delete, upload into) addresses items by ID, so omitting them would
  // force a second lookup for anything beyond just reading the names.
  private async listChildren(path: string): Promise<string> {
    const data = await this.graphRequest(path) as {
      value: Array<{
        id: string;
        name: string;
        folder?: unknown;
        size?: number;
        lastModifiedDateTime: string;
      }>
    };

    if (!data.value?.length) return "📁 Folder is empty.";

    const items = data.value.map((item) => {
      const type = item.folder ? "📁" : "📄";
      const size = item.folder ? "" : ` (${Math.round((item.size ?? 0) / 1024)}KB)`;
      const date = fmtDate(item.lastModifiedDateTime);
      return `${type} ${item.name}${size} — ${date} [ID: ${item.id}]`;
    });

    return `Contents:\n${items.join("\n")}`;
  }

  // ── SharePoint: List folder ───────────────────────────────────────────────
  async listSharePointFolder(
    siteId: string,
    itemId: string,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    const path = itemId === "root"
      ? `${base}/root/children`
      : `${base}/items/${itemId}/children`;
    return this.listChildren(path);
  }

  // ── SharePoint: Upload file ───────────────────────────────────────────────
  async uploadFileToSharePoint(
    siteId: string,
    parentItemId: string,
    fileName: string,
    args: Record<string, unknown>,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    const path = parentItemId === "root"
      ? `${base}/root:/${encodeURI(fileName)}:/content`
      : `${base}/items/${parentItemId}:/${encodeURI(fileName)}:/content`;

    const { body, contentType } = this.resolveUploadBody(args, fileName);
    const data = await this.graphUpload(path, body, contentType) as { webUrl: string };
    log(`Uploaded file to SharePoint: ${fileName} (${contentType})`);
    return `✅ File '${fileName}' uploaded successfully.\nURL: ${data.webUrl}`;
  }

  // ── SharePoint: Move file ─────────────────────────────────────────────────
  async moveSharePointFile(
    siteId: string,
    itemId: string,
    destinationItemId: string,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    await this.graphRequest(`${base}/items/${itemId}`, "PATCH", {
      parentReference: { id: destinationItemId },
    });
    log(`Moved SharePoint item: ${itemId} → ${destinationItemId}`);
    return `✅ Item moved successfully.`;
  }

  // ── SharePoint: Delete file ───────────────────────────────────────────────
  async deleteSharePointFile(
    siteId: string,
    itemId: string,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    await this.graphRequest(`${base}/items/${itemId}`, "DELETE");
    log(`Deleted SharePoint item: ${itemId}`);
    return `✅ Item deleted successfully.`;
  }

  // ── OneDrive: Create folder ───────────────────────────────────────────────
  async createOneDriveFolder(
    parentItemId: string,
    folderName: string
  ): Promise<string> {
    const parentPath = parentItemId === "root"
      ? `/me/drive/root/children`
      : `/me/drive/items/${parentItemId}/children`;

    await this.graphRequest(parentPath, "POST", {
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });

    log(`Created OneDrive folder: ${folderName}`);
    return `✅ Folder '${folderName}' created in OneDrive.`;
  }

  // ── OneDrive: Upload file ─────────────────────────────────────────────────
  async uploadFileToOneDrive(
    parentItemId: string,
    fileName: string,
    args: Record<string, unknown>
  ): Promise<string> {
    const path = parentItemId === "root"
      ? `/me/drive/root:/${encodeURI(fileName)}:/content`
      : `/me/drive/items/${parentItemId}:/${encodeURI(fileName)}:/content`;

    const { body, contentType } = this.resolveUploadBody(args, fileName);
    const data = await this.graphUpload(path, body, contentType) as { webUrl: string };
    log(`Uploaded file to OneDrive: ${fileName} (${contentType})`);
    return `✅ File '${fileName}' uploaded to OneDrive.\nURL: ${data.webUrl}`;
  }

  // ── OneDrive: List folder ─────────────────────────────────────────────────
  // Addressable three ways, cheapest first: nothing at all (root), a path, or
  // an item ID. The path form matters because users name folders, not IDs —
  // without it the model has to list the root and then re-list, one level per
  // round trip, just to reach a folder it was already told the name of.
  async listOneDriveFolder(itemId?: string, folderPath?: string): Promise<string> {
    let path: string;
    const cleanPath = (folderPath ?? "").replace(/^\/+|\/+$/g, "");
    if (cleanPath) {
      // encodeURI (not encodeURIComponent) so '/' separators survive.
      path = `/me/drive/root:/${encodeURI(cleanPath)}:/children`;
    } else if (itemId && itemId !== "root") {
      path = `/me/drive/items/${itemId}/children`;
    } else {
      path = `/me/drive/root/children`;
    }
    return this.listChildren(path);
  }

  // ── OneDrive: Move / rename ───────────────────────────────────────────────
  // Graph expresses both as a PATCH on the item, so one tool covers them and
  // can do both in a single call.
  async moveOneDriveFile(
    itemId: string,
    destinationItemId?: string,
    newName?: string
  ): Promise<string> {
    const body: Record<string, unknown> = {};
    if (destinationItemId) {
      // Graph's move expects the parent's real item ID. "root" is a routing
      // alias in URLs, not an ID, so resolve it to the drive root's actual ID
      // rather than sending a path reference the API accepts inconsistently.
      let parentId = destinationItemId;
      if (destinationItemId === "root") {
        const root = await this.graphRequest(`/me/drive/root`) as { id: string };
        parentId = root.id;
      }
      body.parentReference = { id: parentId };
    }
    if (newName) body.name = newName;

    if (Object.keys(body).length === 0) {
      return "Error: supply destinationItemId to move the item, newName to rename it, or both.";
    }

    await this.graphRequest(`/me/drive/items/${itemId}`, "PATCH", body);
    log(`Updated OneDrive item ${itemId}${destinationItemId ? ` → ${destinationItemId}` : ""}${newName ? ` (renamed to ${newName})` : ""}`);
    if (destinationItemId && newName) return `✅ Item moved and renamed to '${newName}'.`;
    if (newName) return `✅ Item renamed to '${newName}'.`;
    return `✅ Item moved successfully.`;
  }

  // ── Read a text file ──────────────────────────────────────────────────────

  async readTextFile(args: Record<string, unknown>): Promise<string> {
    const ref = String(args.file ?? "").trim();
    if (!ref) return "Error: 'file' is required.";
    const maxChars = Math.max(1000, Number(args.maxChars) || 100_000);

    const item = await fetchDriveItem(this.token, ref, args.siteId as string | undefined);
    const ext = item.name.split(".").pop()?.toLowerCase() ?? "";

    // Point at the right tool rather than returning the mangled text that
    // falls out of decoding a zipped or compressed format as UTF-8.
    const WRONG_TOOL: Record<string, string> = {
      docx: "GetDocumentContent_mcp_WordServer",
      doc:  "GetDocumentContent_mcp_WordServer",
      xlsx: "GetDocumentContent_mcp_ExcelServer",
      xls:  "GetDocumentContent_mcp_ExcelServer",
      pptx: "GetPresentationContent",
      ppt:  "GetPresentationContent",
      pdf:  "ocr_search_and_read",
    };
    if (WRONG_TOOL[ext]) {
      return `Error: '${item.name}' is not a text file. Use ${WRONG_TOOL[ext]} to read it.`;
    }

    // Extension lists never cover everything, so decide on the bytes: a NUL in
    // the opening chunk means binary in every text encoding we would meet here.
    const probe = item.buffer.subarray(0, 4096);
    if (probe.includes(0)) {
      return `Error: '${item.name}' appears to be a binary file, not text. If it is a document, use the tool for its format.`;
    }

    let text = item.buffer.toString("utf8");
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // strip BOM

    const totalChars = text.length;
    const truncated = totalChars > maxChars;
    if (truncated) text = text.slice(0, maxChars);

    log(`Read text file '${item.name}' (${totalChars} chars${truncated ? `, truncated to ${maxChars}` : ""})`);

    // State the truncation and the real size: silently returning a prefix would
    // let the model summarise a fraction of a file as though it were the whole.
    const header = truncated
      ? `${item.name} — showing the first ${maxChars} of ${totalChars} characters (truncated; raise maxChars or ask for a specific part):\n\n`
      : `${item.name} — ${totalChars} characters:\n\n`;
    return header + text;
  }

  // ── Cross-drive transfer ──────────────────────────────────────────────────

  // Both OneDrive and a SharePoint site are just "drives" to Graph, and the
  // /drives/{id} form is the only one that can name two different drives in a
  // single request — which is what makes OneDrive <-> SharePoint possible.
  private async resolveDriveId(siteId?: string): Promise<string> {
    const drive = await this.graphRequest(siteId ? `/sites/${siteId}/drive` : `/me/drive`) as { id: string };
    return drive.id;
  }

  // Accepts a path, an item ID, or 'root', because the model has IDs from a
  // listing but the user speaks in paths.
  private async resolveItem(driveId: string, ref: string): Promise<{ id: string; name: string }> {
    const clean = ref.replace(/^\/+|\/+$/g, "");
    const path =
      !clean || clean === "root"
        ? `/drives/${driveId}/root`
        : clean.includes("/") || clean.includes(".")
          ? `/drives/${driveId}/root:/${encodeURI(clean)}`
          : `/drives/${driveId}/items/${clean}`;
    return await this.graphRequest(path) as { id: string; name: string };
  }

  // Graph's copy is asynchronous: it returns 202 plus a monitor URL rather than
  // the finished item. Waiting for that monitor matters here because a "move"
  // must not delete the source until the copy has genuinely landed.
  private async awaitCopy(monitorUrl: string): Promise<void> {
    const DEADLINE = Date.now() + 120_000;
    while (Date.now() < DEADLINE) {
      await new Promise((r) => setTimeout(r, 1500));
      // The monitor URL is pre-authorised; sending our bearer token is rejected.
      const res = await fetch(monitorUrl);
      if (!res.ok) throw new Error(`copy monitor failed (${res.status})`);
      const body = await res.json() as { status?: string; errorCode?: string };
      if (body.status === "completed") return;
      if (body.status === "failed") throw new Error(`copy failed: ${body.errorCode ?? "unknown error"}`);
    }
    throw new Error("copy did not finish within 2 minutes — it may still complete in the background");
  }

  async transferDriveItem(args: Record<string, unknown>): Promise<string> {
    const source = String(args.source ?? "").trim();
    if (!source) return "Error: 'source' is required.";
    const operation = String(args.operation ?? "copy").toLowerCase() === "move" ? "move" : "copy";
    const newName = args.newName ? String(args.newName) : undefined;

    const srcDriveId = await this.resolveDriveId(args.sourceSiteId as string | undefined);
    const dstDriveId = await this.resolveDriveId(args.destinationSiteId as string | undefined);
    const srcItem = await this.resolveItem(srcDriveId, source);
    const dstFolder = await this.resolveItem(dstDriveId, String(args.destinationFolder ?? "root"));

    // Within one drive a move is a simple synchronous PATCH — no copy, no
    // delete, no window where the file exists twice.
    if (operation === "move" && srcDriveId === dstDriveId) {
      const body: Record<string, unknown> = { parentReference: { id: dstFolder.id } };
      if (newName) body.name = newName;
      await this.graphRequest(`/drives/${srcDriveId}/items/${srcItem.id}`, "PATCH", body);
      log(`Moved '${srcItem.name}' within drive ${srcDriveId}`);
      return `✅ Moved '${newName ?? srcItem.name}' to '${dstFolder.name}'.`;
    }

    const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${srcDriveId}/items/${srcItem.id}/copy`, {
      method: "POST",
      headers: { Authorization: `Bearer ${this.token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        parentReference: { driveId: dstDriveId, id: dstFolder.id },
        ...(newName ? { name: newName } : {}),
      }),
    });
    if (!res.ok && res.status !== 202) {
      throw new Error(`copy request failed (${res.status}): ${await res.text()}`);
    }

    const monitor = res.headers.get("location");
    if (monitor) await this.awaitCopy(monitor);

    if (operation === "move") {
      await this.graphRequest(`/drives/${srcDriveId}/items/${srcItem.id}`, "DELETE");
      log(`Moved '${srcItem.name}' across drives ${srcDriveId} → ${dstDriveId}`);
      return `✅ Moved '${newName ?? srcItem.name}' to '${dstFolder.name}'. The original was deleted from the source.`;
    }

    log(`Copied '${srcItem.name}' to drive ${dstDriveId}`);
    return `✅ Copied '${newName ?? srcItem.name}' to '${dstFolder.name}'. The original is unchanged.`;
  }

  // ── OneDrive: Delete ──────────────────────────────────────────────────────
  async deleteOneDriveFile(itemId: string): Promise<string> {
    await this.graphRequest(`/me/drive/items/${itemId}`, "DELETE");
    log(`Deleted OneDrive item: ${itemId}`);
    return `✅ Item deleted from OneDrive (moved to the recycle bin).`;
  }

  // ── Route tool calls ──────────────────────────────────────────────────────
  async handleToolCall(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    try {
      let result: string;

      switch (toolName) {
        case "create_sharepoint_folder":
          result = await this.createSharePointFolder(
            args.siteId as string,
            args.parentItemId as string,
            args.folderName as string,
            args.driveId as string | undefined
          );
          break;
        case "list_sharepoint_folder":
          result = await this.listSharePointFolder(
            args.siteId as string,
            args.itemId as string,
            args.driveId as string | undefined
          );
          break;
        case "upload_file_to_sharepoint":
          result = await this.uploadFileToSharePoint(
            args.siteId as string,
            args.parentItemId as string,
            args.fileName as string,
            args,
            args.driveId as string | undefined
          );
          break;
        case "move_sharepoint_file":
          result = await this.moveSharePointFile(
            args.siteId as string,
            args.itemId as string,
            args.destinationItemId as string,
            args.driveId as string | undefined
          );
          break;
        case "delete_sharepoint_file":
          result = await this.deleteSharePointFile(
            args.siteId as string,
            args.itemId as string,
            args.driveId as string | undefined
          );
          break;
        case "create_onedrive_folder":
          result = await this.createOneDriveFolder(
            args.parentItemId as string,
            args.folderName as string
          );
          break;
        case "upload_file_to_onedrive":
          result = await this.uploadFileToOneDrive(
            args.parentItemId as string,
            args.fileName as string,
            args
          );
          break;
        case "list_onedrive_folder":
          result = await this.listOneDriveFolder(
            args.itemId as string | undefined,
            args.path as string | undefined
          );
          break;
        case "move_onedrive_file":
          result = await this.moveOneDriveFile(
            args.itemId as string,
            args.destinationItemId as string | undefined,
            args.newName as string | undefined
          );
          break;
        case "delete_onedrive_file":
          result = await this.deleteOneDriveFile(args.itemId as string);
          break;
        case "transfer_drive_item":
          result = await this.transferDriveItem(args);
          break;
        case "read_text_file":
          result = await this.readTextFile(args);
          break;
        default:
          result = `Unknown tool: ${toolName}`;
      }

      const isError = /^(Error|Unknown)/.test(result);
      return { content: [{ type: "text", text: result }], ...(isError ? { isError: true } : {}) };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      log(`SharePoint tool error (${toolName}): ${message}`);
      return { content: [{ type: "text", text: `Error: ${message}` }], isError: true };
    }
  }
}

export const SHAREPOINT_TOOL_NAMES = new Set(SHAREPOINT_TOOLS.map((t) => t.name));

// ── Drive item fetch (for email attachments) ─────────────────────────────────

export interface FetchedDriveItem {
  name: string;
  contentType: string;
  buffer: Buffer;
}

export interface DriveItemLink {
  name: string;
  url: string;
}

/**
 * Creates a sharing link for a drive item, for sending in place of an
 * attachment.
 *
 * Defaults to an organisation-scoped view link: anyone at ABK with the link can
 * open it, but it is not public. Anonymous links are possible in Graph but are
 * not offered here — a mail tool should not be able to publish a company file
 * to the open internet as a side effect of "send this to someone".
 */
export async function createDriveItemLink(
  token: string,
  ref: string,
  siteId?: string,
  type: "view" | "edit" = "view"
): Promise<DriveItemLink> {
  const base = siteId ? `/sites/${siteId}/drive` : "/me/drive";
  const clean = ref.replace(/^\/+|\/+$/g, "");
  const looksLikePath = ref.includes("/") || ref.includes(".");
  const itemPath = looksLikePath
    ? `${base}/root:/${encodeURI(clean)}`
    : `${base}/items/${ref}`;

  const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };

  const metaRes = await fetch(`https://graph.microsoft.com/v1.0${itemPath}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!metaRes.ok) {
    throw new Error(`Could not find '${ref}' (${metaRes.status}): ${await metaRes.text()}`);
  }
  const meta = await metaRes.json() as { name: string; folder?: unknown };

  const res = await fetch(`https://graph.microsoft.com/v1.0${itemPath}/createLink`, {
    method: "POST",
    headers,
    body: JSON.stringify({ type, scope: "organization" }),
  });
  if (!res.ok) {
    throw new Error(`Could not create a link for '${meta.name}' (${res.status}): ${await res.text()}`);
  }
  const body = await res.json() as { link?: { webUrl?: string } };
  const url = body.link?.webUrl;
  if (!url) throw new Error(`Graph returned no link for '${meta.name}'.`);

  log(`Created ${type} link for '${meta.name}'`);
  return { name: meta.name, url };
}

/**
 * Downloads a OneDrive (or SharePoint) item's raw bytes so it can be attached
 * to an email. `ref` is either an item ID or a path relative to the drive root
 * ("/Documents/report.pptx") — the model gets IDs from list_onedrive_folder but
 * users speak in paths, so both are accepted.
 *
 * Metadata is fetched separately from content because the attachment needs the
 * real file name and MIME type; deriving those from the ref would produce
 * "root:" garbage for path refs and nothing at all for ID refs.
 */
export async function fetchDriveItem(
  token: string,
  ref: string,
  siteId?: string
): Promise<FetchedDriveItem> {
  const base = siteId ? `/sites/${siteId}/drive` : "/me/drive";
  const looksLikePath = ref.includes("/") || ref.includes(".");
  const clean = ref.replace(/^\/+|\/+$/g, "");
  const itemPath = looksLikePath
    ? `${base}/root:/${encodeURI(clean)}`
    : `${base}/items/${ref}`;

  const headers = { Authorization: `Bearer ${token}` };

  const metaRes = await fetch(`https://graph.microsoft.com/v1.0${itemPath}`, { headers });
  if (!metaRes.ok) {
    throw new Error(`Could not find '${ref}' (${metaRes.status}): ${await metaRes.text()}`);
  }
  const meta = await metaRes.json() as {
    name: string;
    size?: number;
    file?: { mimeType?: string };
    folder?: unknown;
  };
  if (meta.folder) throw new Error(`'${meta.name}' is a folder — only files can be attached.`);

  const contentRes = await fetch(`https://graph.microsoft.com/v1.0${itemPath}/content`, { headers });
  if (!contentRes.ok) {
    throw new Error(`Could not download '${meta.name}' (${contentRes.status}): ${await contentRes.text()}`);
  }
  const buffer = Buffer.from(await contentRes.arrayBuffer());

  log(`Fetched drive item '${meta.name}' (${Math.round(buffer.length / 1024)}KB)`);
  return {
    name: meta.name,
    contentType: meta.file?.mimeType || "application/octet-stream",
    buffer,
  };
}