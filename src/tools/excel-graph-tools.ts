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
  "UpdateWorkbookRange",
  "AppendRowsToWorksheet",
  "ManageWorksheets",
]);

// A1 column letters: 1 -> A, 26 -> Z, 27 -> AA. Needed to build the target
// address when appending, since the row is computed rather than given.
function columnLetter(n: number): string {
  let s = "";
  let i = Math.max(1, n);
  while (i > 0) {
    const rem = (i - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

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
  {
    name: "UpdateWorkbookRange",
    description:
      "Write values into a specific range of cells in an existing Excel workbook — this is real " +
      "in-place editing: only the cells you name change, the rest of the sheet is untouched. " +
      "Give the range in A1 notation ('B2:D5') and a matching grid of values. Use this to correct " +
      "a figure, fill in a column, or update a table. To add rows at the end of existing data use " +
      "AppendRowsToWorksheet instead, which finds the first free row for you.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: { type: "string", description: "SharePoint site ID (if the file is in SharePoint)" },
        worksheet: { type: "string", description: "Worksheet name, e.g. 'Sheet1'." },
        range: { type: "string", description: "Target range in A1 notation, e.g. 'A1' or 'B2:D5'." },
        values: {
          type: "array",
          description:
            "Rows of cell values; each inner array is one row and must match the range's width " +
            "and height exactly. Numbers stay numeric, strings starting with '=' are treated as " +
            "formulas by Excel.",
          items: { type: "array", items: {} },
        },
      },
      required: ["driveItemId", "worksheet", "range", "values"],
    },
  },
  {
    name: "AppendRowsToWorksheet",
    description:
      "Append rows to the end of the data already in a worksheet, without needing to know where " +
      "that data ends — the first empty row is located automatically. Use this for adding entries " +
      "to a list or log. To overwrite specific cells instead, use UpdateWorkbookRange.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: { type: "string", description: "SharePoint site ID (if the file is in SharePoint)" },
        worksheet: { type: "string", description: "Worksheet name, e.g. 'Sheet1'." },
        values: {
          type: "array",
          description: "Rows to append; each inner array is one row of cell values.",
          items: { type: "array", items: {} },
        },
      },
      required: ["driveItemId", "worksheet", "values"],
    },
  },
  {
    name: "ManageWorksheets",
    description:
      "List, add or delete worksheets (tabs) in an existing Excel workbook. Deleting a worksheet " +
      "removes all of its data — confirm with the user first.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Excel file" },
        siteId: { type: "string", description: "SharePoint site ID (if the file is in SharePoint)" },
        action: {
          type: "string",
          enum: ["list", "add", "delete"],
          description: "'list' (default) returns the worksheet names; 'add' creates one; 'delete' removes one.",
        },
        worksheet: { type: "string", description: "Worksheet name — required for 'add' and 'delete'." },
      },
      required: ["driveItemId"],
    },
  },
];

// ── Handler ───────────────────────────────────────────────────────────────────

export class ExcelGraphToolHandler {
  constructor(private token: string) {}

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
      case "CreateWorkbook":
        return this.createWorkbook(args);
      case "GetDocumentContent_mcp_ExcelServer":
        return this.getContent(args);
      case "CreateComment":
        return this.createComment(args);
      case "ReplyToComment_mcp_ExcelServer":
        return this.replyToComment(args);
      case "UpdateWorkbookRange":
        return this.updateRange(args);
      case "AppendRowsToWorksheet":
        return this.appendRows(args);
      case "ManageWorksheets":
        return this.manageWorksheets(args);
      default:
        return `Unknown Excel tool: ${name}`;
    }
  }

  // ── In-place editing ────────────────────────────────────────────────────────
  //
  // Excel is the one Office format the bridge can edit surgically. Graph exposes
  // a live workbook API addressed by cell range, so a single figure can be
  // changed without touching anything else. Word and PowerPoint have no
  // equivalent — there, "editing" means rewriting the whole file.

  private static asGrid(raw: unknown): unknown[][] | null {
    if (!Array.isArray(raw) || raw.length === 0) return null;
    // Tolerate a single flat row (["a","b"]) as well as a proper grid, since
    // that is the shape a model most often produces for one-row writes.
    const rows = Array.isArray(raw[0]) ? raw as unknown[][] : [raw as unknown[]];
    return rows.every((r) => Array.isArray(r)) ? rows : null;
  }

  private async updateRange(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const sheet = String(args.worksheet ?? "");
    const range = String(args.range ?? "");
    if (!sheet || !range) return "Error: 'worksheet' and 'range' are required.";
    const values = ExcelGraphToolHandler.asGrid(args.values);
    if (!values) return "Error: 'values' must be an array of rows, e.g. [[1,2],[3,4]].";

    const path = `${base}/worksheets/${encodeURIComponent(sheet)}/range(address='${range}')`;
    const res = await gf(this.token, path, {
      method: "PATCH",
      body: JSON.stringify({ values }),
    });
    if (!res.ok) {
      // A shape mismatch is by far the most common failure and Graph's own
      // message for it is opaque, so name the likely cause explicitly.
      const hint = res.status === 400
        ? " The grid must match the range exactly — a 2x3 range needs 2 rows of 3 values."
        : "";
      return `Error updating ${sheet}!${range} (${res.status}): ${JSON.stringify(res.body)}${hint}`;
    }
    const cells = values.reduce((n, r) => n + r.length, 0);
    log(`Updated ${sheet}!${range} (${cells} cells)`);
    return `✅ Updated ${cells} cell(s) in ${sheet}!${range}.`;
  }

  private async appendRows(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const sheet = String(args.worksheet ?? "");
    if (!sheet) return "Error: 'worksheet' is required.";
    const values = ExcelGraphToolHandler.asGrid(args.values);
    if (!values) return "Error: 'values' must be an array of rows, e.g. [[\"Jan\", 100]].";

    // usedRange tells us where the existing data stops; appending below it is
    // what makes this safe to call without first reading the sheet.
    const usedRes = await gf(this.token, `${base}/worksheets/${encodeURIComponent(sheet)}/usedRange`);
    let startRow = 1;
    let width = values[0].length;
    if (usedRes.ok) {
      const used = usedRes.body as { rowIndex?: number; rowCount?: number; columnCount?: number; address?: string };
      // rowIndex is 0-based; +rowCount lands on the first free row, +1 converts to A1's 1-based numbering.
      startRow = (used.rowIndex ?? 0) + (used.rowCount ?? 0) + 1;
      width = Math.max(width, used.columnCount ?? width);
    }

    const endRow = startRow + values.length - 1;
    const endCol = columnLetter(values[0].length);
    const address = `A${startRow}:${endCol}${endRow}`;

    const res = await gf(this.token, `${base}/worksheets/${encodeURIComponent(sheet)}/range(address='${address}')`, {
      method: "PATCH",
      body: JSON.stringify({ values }),
    });
    if (!res.ok) return `Error appending to ${sheet} (${res.status}): ${JSON.stringify(res.body)}`;
    log(`Appended ${values.length} row(s) to ${sheet} at ${address}`);
    return `✅ Appended ${values.length} row(s) to ${sheet}, starting at row ${startRow}.`;
  }

  private async manageWorksheets(args: Record<string, unknown>): Promise<string> {
    const base = excelBase(args);
    const action = String(args.action ?? "list").toLowerCase();
    const sheet = String(args.worksheet ?? "");

    if (action === "list") {
      const res = await gf(this.token, `${base}/worksheets`);
      if (!res.ok) return `Error listing worksheets (${res.status}): ${JSON.stringify(res.body)}`;
      const sheets = ((res.body as { value?: Array<{ name: string }> }).value ?? []).map((s) => s.name);
      return sheets.length ? `Worksheets: ${sheets.join(", ")}` : "Workbook has no worksheets.";
    }

    if (!sheet) return `Error: 'worksheet' is required for action '${action}'.`;

    if (action === "add") {
      const res = await gf(this.token, `${base}/worksheets`, {
        method: "POST",
        body: JSON.stringify({ name: sheet }),
      });
      if (!res.ok) return `Error adding worksheet '${sheet}' (${res.status}): ${JSON.stringify(res.body)}`;
      log(`Added worksheet ${sheet}`);
      return `✅ Worksheet '${sheet}' added.`;
    }

    if (action === "delete") {
      const res = await gf(this.token, `${base}/worksheets/${encodeURIComponent(sheet)}`, { method: "DELETE" });
      if (!res.ok) return `Error deleting worksheet '${sheet}' (${res.status}): ${JSON.stringify(res.body)}`;
      log(`Deleted worksheet ${sheet}`);
      return `✅ Worksheet '${sheet}' deleted, along with its data.`;
    }

    return `Error: unknown action '${action}'. Use 'list', 'add' or 'delete'.`;
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
