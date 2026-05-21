import DocumentIntelligence, {
  getLongRunningPoller,
  isUnexpected,
} from "@azure-rest/ai-document-intelligence";
import { AzureKeyCredential } from "@azure/core-auth";

const DI_ENDPOINT = process.env.AZURE_DI_ENDPOINT ?? "";
const DI_KEY      = process.env.AZURE_DI_KEY ?? "";

export const OCR_TOOL = {
  name: "ocr_search_and_read",
  description:
    "OCR a scanned PDF or image-based document from SharePoint/OneDrive and return its full text. " +
    "PREFERRED: pass drive_id + item_id directly (obtained from list_sharepoint_folder) to skip search entirely. " +
    "FALLBACK: pass query to search by filename/keywords when you don't have the IDs yet.",
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

    const query        = args.query       as string | undefined;
    const hintFilename = args.hint_filename as string | undefined;
    const directDriveId = args.drive_id   as string | undefined;
    const directItemId  = args.item_id    as string | undefined;

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
        filename = hintFilename ?? "document";
      } else {
        const found = await this.searchFile(query!, hintFilename);
        driveId  = found.driveId;
        itemId   = found.itemId;
        filename = found.filename;
      }

      const text = await ocrPdfFromGraph(driveId, itemId, this.accessToken);
      return { content: [{ type: "text", text: `📄 **${filename}**\n\n${text}` }] };
    } catch (err: any) {
      return {
        content: [{ type: "text", text: `OCR error: ${err.message ?? err}` }],
      };
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

export async function ocrPdfFromGraph(
  driveId: string,
  itemId: string,
  accessToken: string
): Promise<string> {
  const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const res = await fetch(downloadUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!res.ok) {
    throw new Error(`Graph download failed: ${res.status} ${res.statusText}`);
  }

  const pdfBuffer = Buffer.from(await res.arrayBuffer());

  const client = DocumentIntelligence(DI_ENDPOINT, new AzureKeyCredential(DI_KEY));

  const initialResponse = await client
    .path("/documentModels/{modelId}:analyze", "prebuilt-layout")
    .post({
      contentType: "application/octet-stream",
      body: pdfBuffer,
    });

  if (isUnexpected(initialResponse)) {
    throw new Error(`Azure DI error: ${initialResponse.body.error?.message}`);
  }

  const poller = getLongRunningPoller(client, initialResponse, { intervalInMs: 1000 });
  const result = await poller.pollUntilDone();

  const doc = result.body as any;
  const lines: string[] = [];

  for (const page of doc.analyzeResult?.pages ?? []) {
    lines.push(`\n=== Page ${page.pageNumber} ===`);
    for (const line of page.lines ?? []) {
      lines.push(line.content);
    }
  }

  for (let i = 0; i < (doc.analyzeResult?.tables ?? []).length; i++) {
    const table = doc.analyzeResult.tables[i];
    lines.push(`\n=== Table ${i + 1} ===`);
    for (const cell of table.cells ?? []) {
      lines.push(`[${cell.rowIndex},${cell.columnIndex}] ${cell.content}`);
    }
  }

  return lines.join("\n");
}