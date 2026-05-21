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

  // ── Route tool calls ──────────────────────────────────────────────────────
  async handleToolCall(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }> }> {
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
        default:
          result = `Unknown tool: ${toolName}`;
      }

      return { content: [{ type: "text", text: result }] };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      log(`SharePoint tool error (${toolName}): ${message}`);
      return { content: [{ type: "text", text: `Error: ${message}` }] };
    }
  }
}

export const SHAREPOINT_TOOL_NAMES = new Set(SHAREPOINT_TOOLS.map((t) => t.name));