import DocumentIntelligence, {
  getLongRunningPoller,
  isUnexpected,
} from "@azure-rest/ai-document-intelligence";
import { AzureKeyCredential } from "@azure/core-auth";

const DI_ENDPOINT = process.env.AZURE_DI_ENDPOINT ?? "";
const DI_KEY      = process.env.AZURE_DI_KEY ?? "";

// ── OneDrive OCR Cache ────────────────────────────────────────────────────────
const OCR_CACHE_FOLDER = "OCR Files";

function log(msg: string) {
  process.stderr.write(`[ocr-tool] ${msg}\n`);
}

/**
 * Builds the cache filename from the original filename and itemId.
 * e.g. "invoice_march.pdf" + itemId → "invoice_march_ocr_a3f8c291.txt"
 */
function buildCacheFilename(originalFilename: string, itemId: string): string {
  const base    = originalFilename.replace(/\.[^.]+$/, "");
  const shortId = itemId.replace(/[^a-zA-Z0-9]/g, "").substring(0, 8);
  return `${base}_ocr_${shortId}.txt`;
}

/**
 * Tries to read a cached OCR result from the user's OneDrive "OCR Files" folder.
 * Returns null on cache miss OR if the cached result has 0 page markers (corrupted).
 */
async function readOcrCache(accessToken: string, cacheFilename: string): Promise<string | null> {
  try {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${OCR_CACHE_FOLDER}/${encodeURIComponent(cacheFilename)}:/content`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!res.ok) return null;
    const text = await res.text();
    // Reject cache entries with no page markers — they are corrupted/empty results
    if (countPages(text) === 0) {
      log(`Cache invalid (0 pages): ${cacheFilename} — treating as miss`);
      return null;
    }
    return text;
  } catch {
    return null;
  }
}

/**
 * Saves OCR result to the user's OneDrive "OCR Files" folder.
 * Auto-creates the folder on first use. Fire-and-forget safe.
 * Throws if text has 0 page markers to prevent caching garbage results.
 */
async function saveOcrCache(accessToken: string, cacheFilename: string, text: string): Promise<void> {
  const pages = countPages(text);
  if (pages === 0) {
    throw new Error(`Refusing to cache OCR result with 0 pages for ${cacheFilename}`);
  }
  log(`Saving cache: ${cacheFilename} (${pages} pages)`);

  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${OCR_CACHE_FOLDER}/${encodeURIComponent(cacheFilename)}:/content`;

  const doUpload = () =>
    fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "text/plain; charset=utf-8",
      },
      body: text,
    });

  let res = await doUpload();

  // If folder doesn't exist yet, create it and retry once
  if (res.status === 404 || res.status === 409) {
    await fetch("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: OCR_CACHE_FOLDER,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }),
    });
    res = await doUpload();
  }

  if (res.ok) {
    log(`Cached OCR result: ${cacheFilename}`);
  } else {
    log(`Failed to save OCR cache: ${res.status} ${res.statusText}`);
  }
}

export const OCR_TOOL = {
  name: "ocr_search_and_read",
  description:
    "OCR a scanned PDF or image-based document from SharePoint/OneDrive and return its full text. " +
    "PREFERRED: pass drive_id + item_id directly (from list_sharepoint_folder) to skip search. " +
    "FALLBACK: pass query to search by filename/keywords. " +
    "Returns ALL pages by default. Use page_from/page_to only if you need a specific range. " +
    "After the first OCR run the full text is cached in OneDrive — subsequent calls are instant. " +
    "For large documents (60-100+ pages) OCR runs in the background (~3-8 min); use check_only=true to poll " +
    "for completion without re-triggering OCR. Use force_refresh=true to discard a bad cache and re-run.",
  inputSchema: {
    type: "object",
    properties: {
      drive_id: {
        type: "string",
        description:
          "DriveId of the file (from list_sharepoint_folder parentReference.driveId). " +
          "When provided together with item_id, search is skipped entirely.",
      },
      item_id: {
        type: "string",
        description:
          "ItemId of the file (from list_sharepoint_folder). " +
          "When provided together with drive_id, search is skipped entirely.",
      },
      query: {
        type: "string",
        description:
          "Search terms to find the file when you don't have drive_id/item_id, " +
          "e.g. 'σύμβαση Smith 2024' or 'invoice March'.",
      },
      hint_filename: {
        type: "string",
        description: "Optional filename hint for the result label, e.g. 'smith_contract.pdf'",
      },
      check_only: {
        type: "boolean",
        description:
          "Set to true to check whether a background OCR job has completed (cache exists) " +
          "WITHOUT re-triggering OCR if it hasn't. Returns status: ready (with page count) or still-in-progress. " +
          "Use this to poll after receiving an 'OCR in progress' response for a large document.",
      },
      force_refresh: {
        type: "boolean",
        description:
          "Set to true to ignore any cached OCR result and re-run Azure Document Intelligence from scratch. " +
          "Use this when the cached result is incomplete or incorrect (e.g. shows fewer pages than expected).",
      },
      page_from: {
        type: "number",
        description: "First page to return (1-based). If omitted, returns from page 1.",
      },
      page_to: {
        type: "number",
        description: "Last page to return (inclusive). If omitted, returns all pages to the end.",
      },
    },
  },
};

export const OCR_TOOL_NAMES = new Set(["ocr_search_and_read"]);

export class OcrToolHandler {
  constructor(private accessToken: string) {}

  async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: { type: string; text: string }[] }> {
    if (name !== "ocr_search_and_read") {
      return { content: [{ type: "text", text: `Unknown OCR tool: ${name}` }] };
    }

    const query         = args.query         as string | undefined;
    const hintFilename  = args.hint_filename as string | undefined;
    const directDriveId = args.drive_id      as string | undefined;
    const directItemId  = args.item_id       as string | undefined;
    const checkOnly     = args.check_only    as boolean | undefined;
    const forceRefresh  = args.force_refresh as boolean | undefined;
    const pageFrom      = args.page_from     as number | undefined;
    const pageTo        = args.page_to       as number | undefined;

    if (!directDriveId && !directItemId && !query) {
      return {
        content: [{ type: "text", text: "OCR error: provide either (drive_id + item_id) or query." }],
      };
    }

    try {
      let driveId: string;
      let itemId: string;
      let filename: string;

      if (directDriveId && directItemId) {
        driveId  = directDriveId;
        itemId   = directItemId;
        filename = hintFilename ?? await this.fetchItemFilename(directDriveId, directItemId);
      } else {
        const found = await this.searchFile(query!, hintFilename);
        driveId  = found.driveId;
        itemId   = found.itemId;
        filename = found.filename;
      }

      const cacheFilename = buildCacheFilename(filename, itemId);

      // ── check_only: poll for background completion without starting OCR ──────
      if (checkOnly) {
        const cached = await readOcrCache(this.accessToken, cacheFilename);
        if (cached) {
          const total = countPages(cached);
          log(`check_only: cache ready — ${total} pages`);
          return {
            content: [{
              type: "text",
              text:
                `✅ **${filename}** — OCR complete (${total} pages cached)\n\n` +
                `Call again without \`check_only\` to read the full document.`,
            }],
          };
        }
        return {
          content: [{
            type: "text",
            text:
              `⏳ **${filename}** — OCR still in progress\n\n` +
              `The cache is not ready yet. Check again in 1–2 minutes.`,
          }],
        };
      }

      // ── 1. Check OneDrive OCR cache ──────────────────────────────────────────
      if (!forceRefresh) {
        const cached = await readOcrCache(this.accessToken, cacheFilename);
        if (cached) {
          log(`Cache hit: ${cacheFilename}`);
          return {
            content: [{
              type: "text",
              text: buildPagedResponse(cached, filename, "*(OCR cache)*", pageFrom, pageTo),
            }],
          };
        }
      } else {
        log(`force_refresh=true: skipping cache for ${cacheFilename}`);
      }

      // ── 2. Cache miss → run Azure Document Intelligence ──────────────────────
      log(`Cache miss: ${cacheFilename} — running OCR...`);
      const onedrivePath = `OCR Files/${cacheFilename}`;

      // Race OCR against a 170 s timeout (stays safely under ACA's 240 s HTTP
      // response limit). For large documents (100+ pages) Azure DI takes 5–8
      // minutes — we fire-and-forget and let the agent check back later via the
      // OneDrive cache. The background promise continues running in the Node.js
      // event loop; ACA keeps the container alive (min_replicas = 1).
      const SYNC_TIMEOUT_MS = 170_000;
      const ocrPromise = ocrPdfFromGraph(driveId, itemId, this.accessToken);

      let fullText: string;
      try {
        fullText = await Promise.race([
          ocrPromise,
          new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error("OCR_TIMEOUT")), SYNC_TIMEOUT_MS)
          ),
        ]);
      } catch (raceErr: any) {
        if (raceErr.message === "OCR_TIMEOUT") {
          // OCR is still running in the background — attach save handler
          ocrPromise
            .then((text) =>
              saveOcrCache(this.accessToken, cacheFilename, text).catch((e) =>
                log(`Background cache save error: ${e}`)
              )
            )
            .catch((e) => log(`Background OCR error: ${e}`));

          return {
            content: [{
              type: "text",
              text:
                `⏳ **${filename}** — OCR in progress (large document)\n\n` +
                `Azure Document Intelligence is analyzing this document in the background. ` +
                `Large scanned PDFs typically take **3–8 minutes** to process.\n\n` +
                `📁 The full text will be saved to OneDrive at:\n` +
                `**\`${onedrivePath}\`**\n\n` +
                `👉 Use \`check_only=true\` (with the same drive_id/item_id) to poll for completion ` +
                `without re-triggering OCR. Check every 1–2 minutes.`,
            }],
          };
        }
        throw raceErr; // re-throw real errors (Azure DI failure, network, etc.)
      }

      // ── 3. OCR completed within timeout — save to OneDrive cache ─────────────
      await saveOcrCache(this.accessToken, cacheFilename, fullText).catch((err) =>
        log(`Cache save error: ${err}`)
      );

      // ── 4. Return requested page range (or first chunk) ──────────────────────
      return {
        content: [{
          type: "text",
          text: buildPagedResponse(fullText, filename, "— OCR complete", pageFrom, pageTo),
        }],
      };

    } catch (err: any) {
      return {
        content: [{ type: "text", text: `OCR error: ${err.message ?? err}` }],
      };
    }
  }

  /**
   * Fetches the actual filename for a known driveId+itemId.
   * Used when the caller has direct IDs but no hint_filename,
   * so the OneDrive OCR cache key is always accurate.
   */
  private async fetchItemFilename(driveId: string, itemId: string): Promise<string> {
    try {
      const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=name`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${this.accessToken}` } });
      if (!res.ok) return "document";
      const data = await res.json() as any;
      return data.name ?? "document";
    } catch {
      return "document";
    }
  }

  private async searchFile(
    query: string,
    hintFilename?: string
  ): Promise<{ driveId: string; itemId: string; filename: string }> {
    const searchQuery = hintFilename ? `${query} ${hintFilename}` : query;

    const body = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: searchQuery },
          fields: ["id", "name", "parentReference", "file"],
          from: 0,
          size: 10,
        },
      ],
    };

    const res = await fetch("https://graph.microsoft.com/v1.0/search/query", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph search failed: ${res.status} ${err}`);
    }

    const data = await res.json() as any;
    const hits: any[] = data?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

    const fileHits = hits.filter((h) => {
      const name: string = h.resource?.name ?? "";
      return /\.(pdf|png|jpg|jpeg|tiff|bmp)$/i.test(name);
    });

    if (fileHits.length === 0) {
      throw new Error(
        `No PDF or image files found matching "${searchQuery}". ` +
        `Try a more specific query or check the filename.`
      );
    }

    const best = fileHits[0].resource;
    const driveId: string = best.parentReference?.driveId;
    const itemId: string = best.id;
    const filename: string = best.name;

    if (!driveId || !itemId) {
      throw new Error(`Could not extract driveId/itemId for "${filename}"`);
    }

    return { driveId, itemId, filename };
  }
}

/**
 * Filters the full OCR text to only include the requested page range.
 * Pages are delimited by "=== Page N ===" markers inserted by ocrPdfFromGraph.
 * Also prepends a page count summary so the agent knows the total.
 */
function countPages(text: string): number {
  return (text.match(/=== Page \d+ ===/g) ?? []).length;
}

function buildPagedResponse(
  text: string,
  filename: string,
  label: string,
  pageFrom?: number,
  pageTo?: number
): string {
  const pageText = applyPageRange(text, pageFrom, pageTo);
  return `📄 **${filename}** ${label}\n\n${pageText}`;
}

function applyPageRange(fullText: string, pageFrom?: number, pageTo?: number): string {
  const pageRegex = /=== Page (\d+) ===/g;
  const pageStarts: { page: number; index: number }[] = [];
  let match: RegExpExecArray | null;

  while ((match = pageRegex.exec(fullText)) !== null) {
    pageStarts.push({ page: parseInt(match[1], 10), index: match.index });
  }

  const totalPages = pageStarts.length;
  const summary = `[📄 Total pages: ${totalPages}${pageFrom ? ` | Showing pages ${pageFrom}–${pageTo ?? totalPages}` : ""}]\n`;

  if (!pageFrom && !pageTo) return summary + fullText;
  if (totalPages === 0)    return summary + fullText;

  const from = pageFrom ?? 1;
  const to   = pageTo   ?? totalPages;

  const startEntry = pageStarts.find((p) => p.page === from);
  const endEntry   = pageStarts.find((p) => p.page === to + 1);

  if (!startEntry) return summary + fullText;
  const slice = endEntry
    ? fullText.slice(startEntry.index, endEntry.index)
    : fullText.slice(startEntry.index);

  return summary + slice;
}

export async function ocrPdfFromGraph(
  driveId: string,
  itemId: string,
  accessToken: string
): Promise<string> {
  const client = DocumentIntelligence(DI_ENDPOINT, new AzureKeyCredential(DI_KEY));

  // ── Always binary mode ────────────────────────────────────────────────────
  // URL mode (@microsoft.graph.downloadUrl) was causing Azure DI to receive
  // only 2 pages for multi-page scanned PDFs — the SharePoint pre-auth URL
  // can return a partial/preview response. Binary mode downloads the full file
  // via Graph API (with Bearer token) and uploads it directly to Azure DI,
  // guaranteeing the complete PDF is processed.
  log(`OCR: downloading ${itemId} via Graph API (binary mode)...`);
  const dlRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!dlRes.ok) {
    throw new Error(`Graph download failed: ${dlRes.status} ${dlRes.statusText}`);
  }
  const pdfBuffer = Buffer.from(await dlRes.arrayBuffer());
  log(`OCR: downloaded ${(pdfBuffer.length / 1024 / 1024).toFixed(1)} MB — submitting to Azure DI...`);

  const initialResponse = await client
    .path("/documentModels/{modelId}:analyze", "prebuilt-layout")
    .post({
      contentType: "application/octet-stream",
      body: pdfBuffer,
    });

  if (isUnexpected(initialResponse)) {
    throw new Error(`Azure DI error: ${initialResponse.body.error?.message}`);
  }

  const poller = getLongRunningPoller(client, initialResponse, { intervalInMs: 2000 });
  const result = await poller.pollUntilDone();

  const doc = result.body as any;
  const analyzeResult = doc.analyzeResult;
  const pages: any[]  = analyzeResult?.pages ?? [];
  const fullContent: string = analyzeResult?.content ?? "";

  log(`Azure DI returned ${pages.length} pages, ${fullContent.length} chars`);

  if (pages.length === 0) {
    throw new Error(
      "Azure Document Intelligence returned 0 pages. " +
      "The PDF may be password-protected, corrupted, or in an unsupported format."
    );
  }

  const lines: string[] = [];

  for (const page of pages) {
    lines.push(`\n=== Page ${page.pageNumber} ===`);

    // PRIMARY: use page.spans to slice the exact text for this page from
    // analyzeResult.content — spans are always present, lines are optional.
    // Offsets are utf16CodeUnit — matches JS string indexing natively.
    if (page.spans?.length > 0 && fullContent) {
      const parts: string[] = [];
      for (const span of page.spans as { offset: number; length: number }[]) {
        parts.push(fullContent.substring(span.offset, span.offset + span.length));
      }
      const pageText = parts.join("").trim();
      if (pageText) { lines.push(pageText); continue; }
    }

    // FALLBACK: explicit lines (may be absent for digital/signed PDFs)
    if (page.lines?.length > 0) {
      for (const line of page.lines) lines.push(line.content);
      continue;
    }

    // LAST RESORT: join individual words
    if (page.words?.length > 0) {
      lines.push((page.words as { content: string }[]).map(w => w.content).join(" "));
    }
  }

  // Append tables (prebuilt-layout only)
  const tables: any[] = analyzeResult?.tables ?? [];
  for (let i = 0; i < tables.length; i++) {
    lines.push(`\n=== Table ${i + 1} ===`);
    for (const cell of tables[i].cells ?? []) {
      lines.push(`[${cell.rowIndex},${cell.columnIndex}] ${cell.content}`);
    }
  }

  return lines.join("\n");
}
