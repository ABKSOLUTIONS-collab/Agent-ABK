/**
 * signature-store.ts
 *
 * Reads the ABK Solutions company logo from OneDrive:
 *   Agent365-Bridge/abk-logo-sig.b64  (base64-encoded PNG)
 *
 * Cached in memory for the lifetime of the process.
 * No Azure Table Storage dependency needed.
 */

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [sig-store] ${msg}\n`);
}

const ONEDRIVE_LOGO_PATH = "Agent365-Bridge/abk-logo-sig.b64";

// ── In-memory cache ───────────────────────────────────────────────────────────
let cachedLogoBase64: string | null = null;

/**
 * Returns the ABK Solutions logo as a base64-encoded PNG string.
 * Reads from OneDrive (Agent365-Bridge/abk-logo-sig.b64).
 * Cached in memory for the lifetime of the process.
 */
export async function getCompanyLogoBase64(graphToken?: string): Promise<string | null> {
  if (cachedLogoBase64) return cachedLogoBase64;
  if (!graphToken) return null;

  try {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_LOGO_PATH}:/content`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${graphToken}` },
    });

    if (resp.status === 404) {
      log("abk-logo-sig.b64 not found in OneDrive");
      return null;
    }
    if (!resp.ok) {
      log(`OneDrive logo fetch failed: ${resp.status}`);
      return null;
    }

    const b64 = (await resp.text()).trim();
    if (!b64) return null;

    cachedLogoBase64 = b64;
    log(`Company logo loaded from OneDrive (${b64.length} chars)`);
    return cachedLogoBase64;
  } catch (err) {
    log(`Error fetching company logo: ${err}`);
    return null;
  }
}
