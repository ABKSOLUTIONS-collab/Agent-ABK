/**
 * excel-graph-tools.ts
 *
 * Excel workbook operations via Microsoft Graph Excel REST API.
 * Works with any standard M365 subscription — no Copilot license required.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { strToU8, zipSync } from "fflate";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [excel-graph] ${msg}\n`);
}

interface GraphFetchResult {
  ok: boolean;
  status: number;
  body: any;
}

// ── Graph fetch helper ────────────────────────────────────────────────────────

async function gf(token: string, path: string, options: RequestInit = {}): Promise<GraphFetchResult> {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers ?? {}),
    },
  });
  const ct = res.headers.get("content-type") ?? "";
  if (ct.includes("json") || ct.includes("odata")) {
    const body = await res.json().catch(() => null);
    return { ok: res.ok, status: res.status, body };
  }
  // Binary or text response
  const text = await res.text().catch(() => "");
  return { ok: res.ok, status: res.status, body: text };
}

// ── Minimal empty XLSX factory ────────────────────────────────────────────────
// Creates the smallest valid Excel workbook with one empty sheet.

function buildEmptyXlsx(sheetName = "Sheet1"): Buffer {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${sheetName}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`;
  const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>ABK Agent</dc:creator>
</cp:coreProperties>`;
  const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Excel</Application>
</Properties>`;

  const files = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rels),
    "xl/workbook.xml": strToU8(workbook),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/worksheets/sheet1.xml": strToU8(sheet),
    "xl/sharedStrings.xml": strToU8(sharedStrings),
    "xl/styles.xml": strToU8(styles),
    "docProps/core.xml": strToU8(core),
    "docProps/app.xml": strToU8(app),
  };
  return Buffer.from(zipSync(files, { level: 6 }));
}

// ── Excel API base path helper ────────────────────────────────────────────────

function excelBase(args: Record<string, unknown>): string {
  const itemId = String(args.driveItemId ?? args.itemId ?? args.workbookId ?? args.id ?? "");
  const siteId = args.siteId as string | undefined;
  if (!itemId) throw new Error("driveItemId is required to identify the Excel workbook.");
  if (siteId) return `/sites/${siteId}/drive/items/${itemId}/workbook`;
  return `/me/drive/items/${itemId}/workbook`;
}

// ── Tool names ────────────────────────────────────────────────────────────────

export const EXCEL_GRAPH_TOOL_NAMES = new Set([
  "CreateWorkbook",
  "GetDocumentContent_mcp_ExcelServer",
  "CreateComment",
  "ReplyToComment_mcp_ExcelServer",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const EXCEL_GRAPH_TOOLS: Tool[] = [
  {
    name: "CreateWorkbook",
    description: "Create a new Excel workbook (.xlsx) in the user's OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "File name (e.g. 'Budget.xlsx')" },
        path: {
          type: "string",
          description: "OneDrive folder path (e.g. '/Documents'). Default: root",
        },
        sheetName: { type: "string", description: "Name for the first worksheet (default: Sheet1)" },
      },
      required: ["name"],
    },
  },
  {
    name: "GetDocumentContent_mcp_ExcelServer",
    description: "Fetch the content of an Excel workbook from OneDrive or SharePoint. Returns worksheet names and cell data.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: {
          type: "string",
          description: "SharePoint site ID (if file is in SharePoint)",
        },
        sheetName: {
          type: "string",
          description: "Specific worksheet to read. If omitted, reads all sheets.",
        },
        maxRows: { type: "number", description: "Max rows per sheet to return (default: 100)" },
      },
    },
  },
  {
    name: "CreateComment",
    description: "Add a comment to a cell in an Excel workbook.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: { type: "string", description: "SharePoint site ID (if file is in SharePoint)" },
        sheetName: {
          type: "string",
          description: "Worksheet name where the comment should be added",
        },
        address: { type: "string", description: "Cell address (e.g. 'A1', 'B3')" },
        content: { type: "string", description: "Comment text" },
      },
      required: ["driveItemId", "content", "address"],
    },
  },
  {
    name: "ReplyToComment_mcp_ExcelServer",
    description: "Reply to an existing comment thread in an Excel workbook.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: { type: "string", description: "SharePoint site ID (if file is in SharePoint)" },
        commentId: { type: "string", description: "ID of the comment to reply to" },
        content: { type: "string", description: "Reply text" },
      },
      required: ["driveItemId", "commentId", "content"],
    },
  },
];

// ── Handler ───────────────────────────────────────────────────────────────────

export class ExcelGraphToolHandler {
  constructor(private token: string) {}

  async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }> }> {
    try {
      const text = await this.dispatch(name, args);
      return { content: [{ type: "text", text }] };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      log(`${name} error: ${msg}`);
      return { content: [{ type: "text", text: `Error: ${msg}` }] };
    }
  }

  private async dispatch(name: string, args: Record<string, unknown>): Promise<string> {
    switch (name) {
      case "CreateWorkbook":
        return this.createWorkbook(args);
      case "GetDocumentContent_mcp_ExcelServer":
        return this.getContent(args);
      case "CreateComment":
        return this.createComment(args);
      case "ReplyToComment_mcp_ExcelServer":
        return this.replyToComment(args);
      default:
        return `Unknown Excel tool: ${name}`;
    }
  }

  // ── Create workbook ─────────────────────────────────────────────────────────

  private async createWorkbook(args: Record<string, unknown>): Promise<string> {
    const rawName = String(args.name ?? "Workbook.xlsx");
    const fileName = rawName.endsWith(".xlsx") ? rawName : rawName + ".xlsx";
    const folder = String(args.path ?? "/").replace(/\/$/, "");
    const sheetName = String(args.sheetName ?? "Sheet1");
    const xlsxBuffer = buildEmptyXlsx(sheetName);

    const uploadPath = folder
      ? `/me/drive/root:${folder}/${encodeURIComponent(fileName)}:/content`
      : `/me/drive/root:/${encodeURIComponent(fileName)}:/content`;

    const res = await gf(this.token, uploadPath, {
      method: "PUT",
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      body: xlsxBuffer,
    });
    if (!res.ok) return `Error creating workbook (${res.status}): ${JSON.stringify(res.body)}`;
    const item = res.body;
    return [
      "Excel workbook created.",
      `Name: ${item.name ?? fileName}`,
      `ID: ${item.id ?? ""}`,
      item.webUrl ? `Open in browser: ${item.webUrl}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }

  // ── Get content ─────────────────────────────────────────────────────────────

  private async getContent(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const maxRows = Number(args.maxRows ?? 100);

    // List worksheets
    const sheetsRes = await gf(this.token, `${base}/worksheets`);
    if (!sheetsRes.ok) return `Error listing worksheets (${sheetsRes.status}): ${JSON.stringify(sheetsRes.body)}`;
    const sheetsData = sheetsRes.body;
    const sheets = sheetsData.value ?? [];
    if (sheets.length === 0) return "Workbook has no worksheets.";

    const targetSheets = args.sheetName ? sheets.filter((s: any) => s.name === args.sheetName) : sheets;
    if (targetSheets.length === 0)
      return `Sheet "${args.sheetName}" not found. Available sheets: ${sheets.map((s: any) => s.name).join(", ")}`;

    const parts: string[] = [];
    for (const sheet of targetSheets) {
      const rangeRes = await gf(this.token, `${base}/worksheets/${encodeURIComponent(sheet.name)}/usedRange`);
      if (!rangeRes.ok) {
        parts.push(`Sheet "${sheet.name}": (error reading data)`);
        continue;
      }
      const rangeData = rangeRes.body;
      const values = rangeData.values ?? [];
      if (values.length === 0) {
        parts.push(`Sheet "${sheet.name}": (empty)`);
        continue;
      }
      const visibleRows = values.slice(0, maxRows);
      const formatted = visibleRows
        .map((row: any[]) => row.map((cell) => String(cell ?? "")).join("\t"))
        .join("\n");
      parts.push(
        `Sheet: ${sheet.name} (${rangeData.rowCount ?? 0} rows × ${rangeData.columnCount ?? 0} cols)\n${formatted}${
          values.length > maxRows ? `\n... (${values.length - maxRows} more rows)` : ""
        }`
      );
    }
    return parts.join("\n\n---\n\n");
  }

  // ── Comments ────────────────────────────────────────────────────────────────

  private async createComment(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const address = String(args.address ?? "A1");
    const content = String(args.content ?? "");
    const sheetName = args.sheetName as string | undefined;

    // Resolve full cell reference (sheet!cell or just cell)
    const cellRef = sheetName ? `${sheetName}!${address}` : address;

    const res = await gf(this.token, `${base}/comments`, {
      method: "POST",
      body: JSON.stringify({
        content: { contentType: "plain", content },
        cellAddress: cellRef,
      }),
    });
    if (!res.ok) return `Error creating comment (${res.status}): ${JSON.stringify(res.body)}`;
    const comment = res.body;
    return `Comment added at ${cellRef}.\nComment ID: ${comment.id ?? ""}`;
  }

  private async replyToComment(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const commentId = String(args.commentId ?? "");
    const content = String(args.content ?? "");

    const res = await gf(this.token, `${base}/comments/${commentId}/replies`, {
      method: "POST",
      body: JSON.stringify({
        content: { contentType: "plain", content },
      }),
    });
    if (!res.ok) return `Error replying to comment (${res.status}): ${JSON.stringify(res.body)}`;
    const reply = res.body;
    return `Reply added to comment ${commentId}.\nReply ID: ${reply.id ?? ""}`;
  }
}
