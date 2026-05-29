/**
 * signature-service.ts
 *
 * Fetches the user's Outlook email signature by reading their most recent
 * sent HTML email via Microsoft Graph (requires Mail.Read scope on the
 * Graph token). The extracted signature is cached in memory per user for
 * 1 hour so we don't hit Graph on every send.
 *
 * Injection: appended to the email body before the agent's email is forwarded
 * to the upstream Agent365 MCP server.
 */

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [signature] ${msg}\n`);
}

// ── In-memory cache ───────────────────────────────────────────────────────────
// Key: first 16 chars of graphToken (enough for isolation, not sensitive)
// Value: { html, fetchedAt }
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

interface CacheEntry {
  html: string;
  fetchedAt: number;
}
const signatureCache = new Map<string, CacheEntry>();

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Returns the user's HTML email signature, or null if it cannot be determined.
 * Results are cached for 1 hour.
 */
export async function getUserSignature(graphToken: string): Promise<string | null> {
  const cacheKey = graphToken.substring(0, 16);
  const cached = signatureCache.get(cacheKey);
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
    log("Using cached signature");
    return cached.html;
  }

  const sig = await fetchSignatureFromSentItems(graphToken);
  if (sig) {
    signatureCache.set(cacheKey, { html: sig, fetchedAt: Date.now() });
    log(`Signature cached (${sig.length} chars)`);
  } else {
    log("No signature found in sent items");
  }
  return sig;
}

/**
 * Appends the HTML signature to an email body.
 * Handles both HTML and plain-text bodies.
 * If the body already contains the signature (idempotency guard), returns unchanged.
 */
export function appendSignature(
  body: string,
  bodyType: string,
  signature: string
): string {
  if (!signature || !body) return body;

  // Idempotency: don't double-append
  // We detect our own injection by looking for a known stable part of the sig
  const sigAnchor = extractIdempotencyAnchor(signature);
  if (sigAnchor && body.includes(sigAnchor)) {
    log("Signature already present — skipping injection");
    return body;
  }

  const isHtml = bodyType?.toLowerCase() === "html";

  if (!isHtml) {
    // Plain-text fallback: strip HTML tags from signature
    const textSig = signature
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
    return `${body}\n\n--\n${textSig}`;
  }

  // HTML: wrap signature in a standard separator div
  const sigBlock = `<div style="margin-top:16px;padding-top:16px;border-top:1px solid #e0e0e0;">${signature}</div>`;

  // Insert before </body> if present; otherwise append
  if (/<\/body>/i.test(body)) {
    return body.replace(/<\/body>/i, `${sigBlock}</body>`);
  }
  return `${body}${sigBlock}`;
}

/**
 * Invalidates the cached signature for a given graph token (e.g., after re-login).
 */
export function invalidateSignatureCache(graphToken: string): void {
  signatureCache.delete(graphToken.substring(0, 16));
}

// ── Graph API fetch ───────────────────────────────────────────────────────────

async function fetchSignatureFromSentItems(graphToken: string): Promise<string | null> {
  try {
    // Fetch the 5 most recent HTML sent emails
    const url =
      "https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages" +
      "?$top=5&$select=body&$orderby=sentDateTime+desc";

    const resp = await fetch(url, {
      headers: {
        Authorization: `Bearer ${graphToken}`,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    });

    if (resp.status === 403 || resp.status === 401) {
      log(`Graph Mail.Read not authorized (${resp.status}) — skipping signature injection`);
      return null;
    }

    if (!resp.ok) {
      log(`Graph sentItems fetch failed: ${resp.status} ${resp.statusText}`);
      return null;
    }

    const data = await resp.json() as {
      value?: Array<{ body?: { contentType?: string; content?: string } }>;
    };

    for (const msg of data.value ?? []) {
      const contentType = msg.body?.contentType ?? "";
      const content = msg.body?.content ?? "";
      if (contentType.toLowerCase() !== "html" || !content) continue;

      const sig = extractSignature(content);
      if (sig) return sig;
    }

    return null;
  } catch (err) {
    log(`Error fetching sent items: ${err}`);
    return null;
  }
}

// ── Signature extraction ──────────────────────────────────────────────────────

/**
 * Attempts to extract the signature block from a full HTML email body.
 * Tries multiple strategies in order of reliability.
 */
function extractSignature(html: string): string | null {
  // Strategy 1: OWA inserts signatures in a <div id="Signature"> element
  const sigById = extractByDivId(html, "Signature");
  if (sigById) { log("Extracted via div#Signature"); return sigById; }

  // Strategy 2: Older Outlook uses <div id="appendonsend">
  const appendOnSend = extractByDivId(html, "appendonsend");
  if (appendOnSend) { log("Extracted via div#appendonsend"); return appendOnSend; }

  // Strategy 3: Look for the divider between main body and signature.
  // OWA's compose wrapper is typically <div id="divtagdefaultwrapper"> (user content)
  // followed immediately by the signature div.
  const afterWrapper = extractAfterMainWrapper(html);
  if (afterWrapper) { log("Extracted after divtagdefaultwrapper"); return afterWrapper; }

  // Strategy 4: Company-specific — find the first div that contains
  // an image whose src/alt contains "abk" (the ABK Solutions logo).
  const byLogo = extractByCompanyLogo(html);
  if (byLogo) { log("Extracted via ABK logo anchor"); return byLogo; }

  // Strategy 5: Last resort — extract content after the last <hr> tag
  const afterHr = extractAfterLastHr(html);
  if (afterHr) { log("Extracted after <hr>"); return afterHr; }

  return null;
}

/** Extracts a <div id="..."> and its content (handles nested divs). */
function extractByDivId(html: string, id: string): string | null {
  const re = new RegExp(`<div[^>]+id=["']${id}["'][^>]*>`, "i");
  const match = re.exec(html);
  if (!match) return null;

  const start = match.index;
  const innerStart = start + match[0].length;

  // Walk forward counting open/close divs to find the matching </div>
  let depth = 1;
  let i = innerStart;
  while (i < html.length && depth > 0) {
    const openIdx = html.indexOf("<div", i);
    const closeIdx = html.indexOf("</div>", i);
    if (closeIdx < 0) break;
    if (openIdx >= 0 && openIdx < closeIdx) {
      depth++;
      i = openIdx + 4;
    } else {
      depth--;
      if (depth === 0) {
        return html.substring(start, closeIdx + 6).trim() || null;
      }
      i = closeIdx + 6;
    }
  }
  return null;
}

/** Extracts content that follows OWA's main content wrapper div. */
function extractAfterMainWrapper(html: string): string | null {
  const wrapperRe = /<div[^>]+id=["']divtagdefaultwrapper["'][^>]*>/i;
  const wrapperMatch = wrapperRe.exec(html);
  if (!wrapperMatch) return null;

  // Find the end of this div
  let depth = 1;
  let i = wrapperMatch.index + wrapperMatch[0].length;
  while (i < html.length && depth > 0) {
    const openIdx = html.indexOf("<div", i);
    const closeIdx = html.indexOf("</div>", i);
    if (closeIdx < 0) break;
    if (openIdx >= 0 && openIdx < closeIdx) {
      depth++;
      i = openIdx + 4;
    } else {
      depth--;
      i = closeIdx + 6;
    }
  }

  // Everything after the wrapper div, up to </body>
  const after = html.substring(i).replace(/<\/body>[\s\S]*$/i, "").trim();

  // Only return if there's substantial content (not just whitespace/empty divs)
  const textContent = after.replace(/<[^>]+>/g, "").trim();
  if (textContent.length > 20) return after;
  return null;
}

/** Finds the first top-level div containing an image with "abk" in its src/alt. */
function extractByCompanyLogo(html: string): string | null {
  const imgRe = /<img[^>]+(src|alt)=["'][^"']*abk[^"']*["'][^>]*>/i;
  const imgMatch = imgRe.exec(html);
  if (!imgMatch) return null;

  // Walk backward to find the containing top-level div
  const beforeImg = html.substring(0, imgMatch.index);
  const lastDivStart = beforeImg.lastIndexOf("<div");
  if (lastDivStart < 0) return null;

  // Extract from that div to </body> (or end of content)
  const fromDiv = html.substring(lastDivStart);
  const bodyClose = fromDiv.search(/<\/body>/i);
  const snippet = bodyClose >= 0 ? fromDiv.substring(0, bodyClose) : fromDiv;
  return snippet.trim() || null;
}

/** Extracts HTML after the last <hr> tag in the email. */
function extractAfterLastHr(html: string): string | null {
  const hrIdx = html.lastIndexOf("<hr");
  if (hrIdx < 0) return null;

  const after = html.substring(hrIdx).replace(/<\/body>[\s\S]*$/i, "").trim();
  const textContent = after.replace(/<[^>]+>/g, "").trim();
  if (textContent.length > 20) return after;
  return null;
}

/**
 * Picks a short, stable string from the signature to use as an idempotency
 * anchor — prevents double-injection if the agent somehow calls send twice.
 */
function extractIdempotencyAnchor(signature: string): string | null {
  // Use the first 40 chars of text content as the anchor
  const text = signature.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
  if (text.length >= 20) return text.substring(0, 40);
  return null;
}
