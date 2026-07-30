/**
 * word-graph-tools.ts
 *
 * Word document operations via Microsoft Graph + DOCX manipulation.
 * Works with any standard M365 subscription — no Copilot license required.
 *
 * Uses fflate for ZIP (DOCX) read/write.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { strToU8, strFromU8, zipSync, unzipSync } from "fflate";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [word-graph] ${msg}\n`);
}

interface GraphFetchResult {
  ok: boolean;
  status: number;
  body: any;
  rawBuffer?: Buffer;
}

// ── Graph fetch helper ────────────────────────────────────────────────────────

async function gf(token: string, path: string, options: RequestInit = {}): Promise<GraphFetchResult> {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers ?? {}),
    },
  });
  const ct = res.headers.get("content-type") ?? "";
  if (ct.includes("json") || ct.includes("odata")) {
    const body = await res.json().catch(() => null);
    return { ok: res.ok, status: res.status, body };
  }
  const ab = await res.arrayBuffer();
  const rawBuffer = Buffer.from(ab);
  return { ok: res.ok, status: res.status, body: rawBuffer, rawBuffer };
}

// ── XML helpers ───────────────────────────────────────────────────────────────

function xmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function wrapParagraphs(text: string): string {
  return text
    .split("\n")
    .map((line) => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(line)}</w:t></w:r></w:p>`)
    .join("\n");
}

// ── DOCX XML templates ────────────────────────────────────────────────────────

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
</Properties>`;

function coreXml(title: string): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${xmlEscape(title)}</dc:title>
  <dc:creator>ABK Agent</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
</cp:coreProperties>`;
}

function documentXml(bodyContent: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyContent}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function emptyCommentsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
}

// ── DOCX factory ──────────────────────────────────────────────────────────────

function buildDocx(title: string, bodyText: string): Buffer {
  const paragraphs = wrapParagraphs(bodyText || " ");
  const files = {
    "[Content_Types].xml": strToU8(CONTENT_TYPES_XML),
    "_rels/.rels": strToU8(RELS_XML),
    "word/document.xml": strToU8(documentXml(paragraphs)),
    "word/_rels/document.xml.rels": strToU8(DOCUMENT_RELS_XML),
    "word/settings.xml": strToU8(SETTINGS_XML),
    "word/comments.xml": strToU8(emptyCommentsXml()),
    "docProps/core.xml": strToU8(coreXml(title)),
    "docProps/app.xml": strToU8(APP_XML),
  };
  return Buffer.from(zipSync(files, { level: 6 }));
}

// ── DOCX text extractor ───────────────────────────────────────────────────────

function extractDocxText(buffer: Buffer): string {
  try {
    const files = unzipSync(buffer);
    const docXml = files["word/document.xml"];
    if (!docXml) return "(Could not find document.xml inside the DOCX)";
    const xml = strFromU8(docXml);

    // Extract paragraph text, preserving paragraph breaks
    const paragraphs: string[] = [];
    let current: string[] = [];
    for (const match of xml.matchAll(/<(w:p|w:br|w:t)(?:\s[^>]*)?>([^<]*)/g)) {
      const tag = match[1];
      const text = match[2];
      if (tag === "w:p") {
        if (current.length > 0) paragraphs.push(current.join(""));
        current = [];
      } else if (tag === "w:t" && text) {
        current.push(text);
      }
    }
    if (current.length > 0) paragraphs.push(current.join(""));

    const result = paragraphs.filter(Boolean).join("\n").replace(/\s+/g, " ").trim();
    return result || "(Document is empty)";
  } catch (err) {
    return `(Error extracting text: ${err instanceof Error ? err.message : String(err)})`;
  }
}

// ── DOCX comment adder ────────────────────────────────────────────────────────

function addCommentToDocx(buffer: Buffer, author: string, commentText: string): Buffer {
  const files = unzipSync(buffer);

  // Parse existing comments.xml or start fresh
  const existingCommentsXml = files["word/comments.xml"] ? strFromU8(files["word/comments.xml"]) : emptyCommentsXml();

  // Count existing comments to get next ID
  const existingIds = Array.from(existingCommentsXml.matchAll(/w:id="(\d+)"/g)).map((m) => parseInt(m[1], 10));
  const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 1;

  const now = new Date().toISOString();
  const initials = author
    .split(/\s+/)
    .map((w) => w[0] ?? "")
    .join("")
    .toUpperCase()
    .substring(0, 3);

  // Build new comment entry
  const newComment = `<w:comment w:id="${nextId}" w:author="${xmlEscape(author)}" w:date="${now}" w:initials="${xmlEscape(initials)}">
    <w:p><w:r><w:t>${xmlEscape(commentText)}</w:t></w:r></w:p>
  </w:comment>`;

  // Inject before closing </w:comments>
  const updatedCommentsXml = existingCommentsXml.includes("</w:comments>")
    ? existingCommentsXml.replace("</w:comments>", `${newComment}\n</w:comments>`)
    : existingCommentsXml + newComment;

  // Anchor the comment to the last paragraph in document.xml
  const docXml = strFromU8(files["word/document.xml"]);
  const commentAnchor = `<w:commentRangeStart w:id="${nextId}"/><w:commentRangeEnd w:id="${nextId}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${nextId}"/></w:r>`;
  // Find last </w:p> before </w:body> and inject the anchor before it
  const updatedDocXml = docXml.replace(/(<\/w:p>)(\s*<w:sectPr)/, `${commentAnchor}$1$2`);

  // Ensure [Content_Types].xml includes comments
  const ctXml = files["[Content_Types].xml"] ? strFromU8(files["[Content_Types].xml"]) : CONTENT_TYPES_XML;
  const updatedCt = ctXml.includes("wordprocessingml.comments")
    ? ctXml
    : ctXml.replace(
        "</Types>",
        `  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>\n</Types>`
      );

  // Ensure word/_rels/document.xml.rels includes comments relationship
  const relsXml = files["word/_rels/document.xml.rels"]
    ? strFromU8(files["word/_rels/document.xml.rels"])
    : DOCUMENT_RELS_XML;
  const updatedRels = relsXml.includes("relationships/comments")
    ? relsXml
    : relsXml.replace(
        "</Relationships>",
        `  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>\n</Relationships>`
      );

  // Rebuild ZIP
  const updatedFiles: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(files)) {
    updatedFiles[name] = data;
  }
  updatedFiles["word/comments.xml"] = strToU8(updatedCommentsXml);
  updatedFiles["word/document.xml"] = strToU8(updatedDocXml);
  updatedFiles["[Content_Types].xml"] = strToU8(updatedCt);
  updatedFiles["word/_rels/document.xml.rels"] = strToU8(updatedRels);

  return Buffer.from(zipSync(updatedFiles, { level: 6 }));
}

// ── Tool names ────────────────────────────────────────────────────────────────

export const WORD_GRAPH_TOOL_NAMES = new Set([
  "CreateDocument",
  "GetDocumentContent_mcp_WordServer",
  "AddComment",
  "ReplyToComment_mcp_WordServer",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const WORD_GRAPH_TOOLS: Tool[] = [
  {
    name: "CreateDocument",
    description: "Create a new Word document (.docx) in the user's OneDrive and optionally populate it with text content.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Document file name (e.g. 'Meeting Notes.docx')" },
        path: {
          type: "string",
          description: "OneDrive folder path to create the document in (e.g. '/Documents'). Default: root",
        },
        content: {
          type: "string",
          description: "Initial text content for the document. Use \\n for paragraph breaks.",
        },
        title: { type: "string", description: "Document title (metadata). Defaults to file name." },
      },
      required: ["name"],
    },
  },
  {
    name: "GetDocumentContent_mcp_WordServer",
    description: "Fetch and return the text content of a Word document (.docx) from OneDrive or SharePoint.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: {
          type: "string",
          description: "OneDrive item ID of the Word document",
        },
        siteId: {
          type: "string",
          description: "SharePoint site ID (required if the file is in SharePoint, not OneDrive)",
        },
        filePath: {
          type: "string",
          description: "OneDrive file path (e.g. '/Documents/Report.docx'). Alternative to driveItemId.",
        },
      },
    },
  },
  {
    name: "AddComment",
    description: "Add a review comment to a Word document. The comment appears in Word's review pane.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: {
          type: "string",
          description: "OneDrive item ID of the Word document",
        },
        siteId: {
          type: "string",
          description: "SharePoint site ID (if file is in SharePoint)",
        },
        comment: { type: "string", description: "Comment text to add" },
        author: { type: "string", description: "Comment author name (defaults to account name)" },
      },
      required: ["comment"],
    },
  },
  {
    name: "ReplyToComment_mcp_WordServer",
    description:
      "Add a new comment (reply note) to a Word document. Note: reply threading is not supported via API — this adds a new top-level comment.",
    inputSchema: {
      type: "object",
      properties: {
        driveItemId: { type: "string", description: "OneDrive item ID of the Word document" },
        siteId: { type: "string", description: "SharePoint site ID (if file is in SharePoint)" },
        comment: { type: "string", description: "Reply/comment text" },
        author: { type: "string", description: "Comment author name" },
      },
      required: ["comment"],
    },
  },
];

// ── Handler ───────────────────────────────────────────────────────────────────

export class WordGraphToolHandler {
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
      case "CreateDocument":
        return this.createDocument(args);
      case "GetDocumentContent_mcp_WordServer":
        return this.getContent(args);
      case "AddComment":
      case "ReplyToComment_mcp_WordServer":
        return this.addComment(args);
      default:
        return `Unknown Word tool: ${name}`;
    }
  }

  // ── Create ──────────────────────────────────────────────────────────────────

  private async createDocument(args: Record<string, unknown>): Promise<string> {
    const rawName = String(args.name ?? "Document.docx");
    const fileName = rawName.endsWith(".docx") ? rawName : rawName + ".docx";
    const folder = String(args.path ?? "/").replace(/\/$/, "");
    const title = String(args.title ?? rawName.replace(/\.docx$/i, ""));
    const content = String(args.content ?? "");
    const docxBuffer = buildDocx(title, content);

    // Upload to OneDrive at the given path
    const uploadPath = folder
      ? `/me/drive/root:${folder}/${encodeURIComponent(fileName)}:/content`
      : `/me/drive/root:/${encodeURIComponent(fileName)}:/content`;

    const res = await gf(this.token, uploadPath, {
      method: "PUT",
      headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      body: docxBuffer,
    });
    if (!res.ok) return `Error creating document (${res.status}): ${JSON.stringify(res.body)}`;
    const item = res.body;
    return [
      "Word document created.",
      `Name: ${item.name ?? fileName}`,
      `ID: ${item.id ?? ""}`,
      item.webUrl ? `Open in browser: ${item.webUrl}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }

  // ── Get content ─────────────────────────────────────────────────────────────

  private async getContent(args: Record<string, unknown>): Promise<string> {
    const { path: drivePath, itemId } = this.resolveItem(args);
    const res = await gf(this.token, drivePath + "/content");
    if (!res.ok) return `Error downloading document (${res.status}): ${JSON.stringify(res.body)}`;
    if (!res.rawBuffer) return "Error: empty response from Graph";
    log(`Downloaded DOCX: ${res.rawBuffer.length} bytes (itemId=${itemId})`);
    const text = extractDocxText(res.rawBuffer);
    return `Document content:\n\n${text}`;
  }

  // ── Add comment ─────────────────────────────────────────────────────────────

  private async addComment(args: Record<string, unknown>): Promise<string> {
    const { path: drivePath } = this.resolveItem(args);
    const commentText = String(args.comment ?? "");
    const author = String(args.author ?? "ABK Agent");

    // Download current DOCX
    const dlRes = await gf(this.token, drivePath + "/content");
    if (!dlRes.ok) return `Error downloading document (${dlRes.status}): ${JSON.stringify(dlRes.body)}`;
    if (!dlRes.rawBuffer) return "Error: empty response from Graph";

    // Add comment to DOCX
    const updatedBuffer = addCommentToDocx(dlRes.rawBuffer, author, commentText);

    // Upload back
    const upRes = await gf(this.token, drivePath + "/content", {
      method: "PUT",
      headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      body: updatedBuffer,
    });
    if (!upRes.ok) return `Error uploading document (${upRes.status}): ${JSON.stringify(upRes.body)}`;
    return `Comment added by "${author}":\n"${commentText}"\n\nThe comment is visible in Word's Review pane.`;
  }

  // ── Resolve item path ────────────────────────────────────────────────────────

  private resolveItem(args: Record<string, unknown>): { path: string; itemId: string } {
    const itemId = String(args.driveItemId ?? args.itemId ?? args.id ?? "");
    const siteId = args.siteId as string | undefined;
    const filePath = args.filePath as string | undefined;

    if (siteId && itemId) {
      return { path: `/sites/${siteId}/drive/items/${itemId}`, itemId };
    }
    if (itemId) {
      return { path: `/me/drive/items/${itemId}`, itemId };
    }
    if (filePath) {
      const encoded = filePath.startsWith("/") ? filePath : `/${filePath}`;
      return { path: `/me/drive/root:${encoded}:`, itemId: filePath };
    }
    throw new Error("Please provide driveItemId or filePath to identify the document.");
  }
}
