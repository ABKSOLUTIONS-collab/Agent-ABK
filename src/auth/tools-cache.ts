import { TableClient, AzureNamedKeyCredential, TableEntity } from "@azure/data-tables";

/**
 * Caches the discovered tool list in Azure Table Storage so that it
 * survives container restarts, deploys, and volume changes.
 */

const ACCOUNT     = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY = process.env.AZURE_STORAGE_KEY ?? "";
const TABLE       = process.env.TOKEN_TABLE_NAME || "agent365tokens";

// Cache row lives in a separate partition so it never conflicts with tokens
const CACHE_PARTITION = "toolscache";
const CACHE_ROW       = "latest";

/** Max age of the cache before it's considered stale (30 days) */
const MAX_CACHE_AGE_MS = 30 * 24 * 60 * 60 * 1000;

export interface CachedTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  serverName: string;
}

interface CacheEntity extends TableEntity {
  timestamp: number;
  toolsJson: string;
}

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

function getClient(): TableClient | null {
  if (!ACCOUNT || !ACCOUNT_KEY) return null;
  const cred = new AzureNamedKeyCredential(ACCOUNT, ACCOUNT_KEY);
  return new TableClient(
    `https://${ACCOUNT}.table.core.windows.net`,
    TABLE,
    cred
  );
}

/**
 * Saves the discovered tool list to Azure Table Storage.
 */
export async function saveToolsCache(tools: CachedTool[]): Promise<void> {
  const client = getClient();
  if (!client) {
    log("saveToolsCache: no storage credentials, skipping");
    return;
  }
  try {
    const entity: CacheEntity = {
      partitionKey: CACHE_PARTITION,
      rowKey:       CACHE_ROW,
      timestamp:    Date.now(),
      toolsJson:    JSON.stringify(tools),
    };
    await client.upsertEntity(entity, "Replace");
    log(`Cached ${tools.length} tools to Table Storage`);
  } catch (e) {
    log(`saveToolsCache error: ${e}`);
  }
}

/**
 * Loads the cached tool list from Azure Table Storage.
 * Returns null if no cache exists or if it's too old.
 */
export async function loadToolsCache(): Promise<CachedTool[] | null> {
  const client = getClient();
  if (!client) {
    log("loadToolsCache: no storage credentials, skipping");
    return null;
  }
  try {
    const entity = await client.getEntity<CacheEntity>(CACHE_PARTITION, CACHE_ROW);
    const age = Date.now() - entity.timestamp;
    if (age > MAX_CACHE_AGE_MS) {
      log(`Tool cache is stale (${Math.round(age / 86400000)}d old), ignoring`);
      return null;
    }
    const tools: CachedTool[] = JSON.parse(entity.toolsJson);
    log(`Loaded ${tools.length} cached tools from Table Storage (${Math.round(age / 60000)}m old)`);
    return tools;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) {
      log(`loadToolsCache error: ${err}`);
    }
    return null;
  }
}

/**
 * Clears the cached tool list from Azure Table Storage.
 */
export async function clearToolsCache(): Promise<void> {
  const client = getClient();
  if (!client) return;
  try {
    await client.deleteEntity(CACHE_PARTITION, CACHE_ROW);
    log("Tool cache cleared from Table Storage");
  } catch { /* already gone */ }
}