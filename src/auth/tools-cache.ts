import { BlobServiceClient, StorageSharedKeyCredential } from "@azure/storage-blob";

/**
 * Caches the discovered tool list in Azure Blob Storage.
 * Blob Storage has no size limit (unlike Table Storage which caps at 64KB),
 * making it suitable for large tool lists (63+ tools).
 */

const ACCOUNT     = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY = process.env.AZURE_STORAGE_KEY ?? "";
const CONTAINER   = "tools-cache";
const BLOB_NAME   = "tools-latest.json";

/** Max age of the cache before it's considered stale (30 days) */
const MAX_CACHE_AGE_MS = 30 * 24 * 60 * 60 * 1000;

export interface CachedTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  serverName: string;
}

interface CachePayload {
  timestamp: number;
  tools: CachedTool[];
}

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

function getContainerClient() {
  if (!ACCOUNT || !ACCOUNT_KEY) return null;
  const cred = new StorageSharedKeyCredential(ACCOUNT, ACCOUNT_KEY);
  const blobService = new BlobServiceClient(
    `https://${ACCOUNT}.blob.core.windows.net`,
    cred
  );
  return blobService.getContainerClient(CONTAINER);
}

/**
 * Saves the discovered tool list to Azure Blob Storage.
 */
export async function saveToolsCache(tools: CachedTool[]): Promise<void> {
  const container = getContainerClient();
  if (!container) {
    log("saveToolsCache: no storage credentials, skipping");
    return;
  }
  try {
    const payload: CachePayload = { timestamp: Date.now(), tools };
    const json = JSON.stringify(payload);
    const blockBlob = container.getBlockBlobClient(BLOB_NAME);
    await blockBlob.upload(json, Buffer.byteLength(json), {
      blobHTTPHeaders: { blobContentType: "application/json" },
    });
    log(`Cached ${tools.length} tools to Blob Storage (${Math.round(json.length / 1024)}KB)`);
  } catch (e) {
    log(`saveToolsCache error: ${e}`);
  }
}

/**
 * Loads the cached tool list from Azure Blob Storage.
 * Returns null if no cache exists or if it's too old.
 */
export async function loadToolsCache(): Promise<CachedTool[] | null> {
  const container = getContainerClient();
  if (!container) {
    log("loadToolsCache: no storage credentials, skipping");
    return null;
  }
  try {
    const blockBlob = container.getBlockBlobClient(BLOB_NAME);
    const download = await blockBlob.download(0);
    const chunks: Buffer[] = [];
    for await (const chunk of download.readableStreamBody as AsyncIterable<Buffer>) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    const json = Buffer.concat(chunks).toString("utf8");
    const payload: CachePayload = JSON.parse(json);
    const age = Date.now() - payload.timestamp;
    if (age > MAX_CACHE_AGE_MS) {
      log(`Tool cache is stale (${Math.round(age / 86400000)}d old), ignoring`);
      return null;
    }
    log(`Loaded ${payload.tools.length} cached tools from Blob Storage (${Math.round(age / 60000)}m old)`);
    return payload.tools;
  } catch (err: unknown) {
    const code = (err as { statusCode?: number; code?: string }).statusCode;
    if (code !== 404) log(`loadToolsCache error: ${err}`);
    return null;
  }
}

/**
 * Clears the cached tool list from Azure Blob Storage.
 */
export async function clearToolsCache(): Promise<void> {
  const container = getContainerClient();
  if (!container) return;
  try {
    await container.getBlockBlobClient(BLOB_NAME).delete();
    log("Tool cache cleared from Blob Storage");
  } catch { /* already gone */ }
}
