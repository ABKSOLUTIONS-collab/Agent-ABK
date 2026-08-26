import { Tool } from "@modelcontextprotocol/sdk/types.js";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
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
    description: "Uploads a text file to a SharePoint folder by item ID.",
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
          description: "File name including extension e.g. 'report.txt'.",
        },
        content: {
          type: "string",
          description: "Text content of the file.",
        },
        driveId: {
          type: "string",
          description: "Optional. Drive ID for a specific document library.",
        },
      },
      required: ["siteId", "parentItemId", "fileName", "content"],
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
    description: "Uploads a text file to the user's OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        parentItemId: {
          type: "string",
          description: "Item ID of the destination folder. Use 'root' for OneDrive root.",
        },
        fileName: {
          type: "string",
          description: "File name including extension e.g. 'notes.txt'.",
        },
        content: {
          type: "string",
          description: "Text content of the file.",
        },
      },
      required: ["parentItemId", "fileName", "content"],
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

  private async graphUpload(path: string, content: string): Promise<unknown> {
    const url = `https://graph.microsoft.com/v1.0${path}`;
    const res = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "text/plain",
      },
      body: content,
    });
    const text = await res.text();
    if (!res.ok) {
      throw new Error(`Graph API upload ${res.status}: ${text}`);
    }
    return JSON.parse(text);
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
      const date = new Date(item.lastModifiedDateTime).toLocaleDateString();
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
    content: string,
    driveId?: string
  ): Promise<string> {
    const base = this.siteDrivePath(siteId, driveId);
    const path = parentItemId === "root"
      ? `${base}/root:/${fileName}:/content`
      : `${base}/items/${parentItemId}:/${fileName}:/content`;

    const data = await this.graphUpload(path, content) as { webUrl: string };
    log(`Uploaded file to SharePoint: ${fileName}`);
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
    content: string
  ): Promise<string> {
    const path = parentItemId === "root"
      ? `/me/drive/root:/${fileName}:/content`
      : `/me/drive/items/${parentItemId}:/${fileName}:/content`;

    const data = await this.graphUpload(path, content) as { webUrl: string };
    log(`Uploaded file to OneDrive: ${fileName}`);
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
            args.content as string,
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
            args.content as string
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

  log(`Fetched drive item '${meta.name}' (${Math.round(buffer.length / 1024)}KB) for attachment`);
  return {
    name: meta.name,
    contentType: meta.file?.mimeType || "application/octet-stream",
    buffer,
  };
}