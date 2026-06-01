/**
 * signature-store.ts
 *
 * Reads the ABK Solutions company logo from Azure Table Storage (table: agent365config).
 * The logo was uploaded once by the admin as a base64-encoded PNG string.
 *
 * User signature TEXT is stored separately in OneDrive (signature-style-tool.ts).
 */

import { TableClient, AzureNamedKeyCredential } from "@azure/data-tables";

const ACCOUNT      = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY  = process.env.AZURE_STORAGE_KEY ?? "";
const CONFIG_TABLE = "agent365config";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [sig-store] ${msg}\n`);
}

function getConfigClient(): TableClient {
  if (!ACCOUNT || !ACCOUNT_KEY) throw new Error("Storage credentials missing");
  return new TableClient(
    `https://${ACCOUNT}.table.core.windows.net`,
    CONFIG_TABLE,
    new AzureNamedKeyCredential(ACCOUNT, ACCOUNT_KEY)
  );
}

// ── In-memory cache ───────────────────────────────────────────────────────────
let cachedLogoBase64: string | null = null;

/**
 * Returns the ABK Solutions logo as a base64-encoded PNG string.
 * Cached in memory for the lifetime of the process.
 */
export async function getCompanyLogoBase64(): Promise<string | null> {
  if (cachedLogoBase64) return cachedLogoBase64;
  try {
    const client = getConfigClient();
    const entity = await client.getEntity<{ value: string }>("config", "abk-logo-png");
    cachedLogoBase64 = entity.value ?? null;
    if (cachedLogoBase64) log("Company logo loaded from Table Storage");
    return cachedLogoBase64;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) log(`Error fetching company logo: ${err}`);
    return null;
  }
}
