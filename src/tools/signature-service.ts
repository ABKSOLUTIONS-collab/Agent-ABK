/**
 * signature-service.ts
 *
 * Provides the user's HTML email signature for auto-injection into outgoing emails.
 *
 * Priority:
 *   1. OneDrive: Agent365-Bridge/signature.txt  (structured, per-user, always up-to-date)
 *   2. Fallback: most recent HTML sent email via Graph (extracts Outlook signature)
 *
 * appendSignature() always produces an HTML body — plain text bodies are
 * converted to HTML so the signature is never rendered as a squished string.
 */

import * as fs from "fs";
import * as path from "path";

const LOGO_CID = "abk-logo@abk";

// Read logo once at startup from local filesystem (bundled in Docker image)
function loadLogoBase64(): string | null {
  try {
    const logoPath = path.join(__dirname, "../../assets/abk-logo-email.png");
    const data = fs.readFileSync(logoPath);
    return data.toString("base64");
  } catch {
    return null;
  }
}

/** Base64 PNG of the ABK Solutions logo — exported for CID attachment injection */
export const LOGO_BASE64: string | null = loadLogoBase64();

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [signature] ${msg}\n`);
}

// ── Cache ─────────────────────────────────────────────────────────────────────
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

interface CacheEntry { html: string; fetchedAt: number; }
const signatureCache = new Map<string, CacheEntry>();

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Returns the user's HTML email signature.
 * Tries OneDrive signature.txt first, then sent items. Cached 1 hour.
 */
export async function getUserSignature(graphToken: string): Promise<string | null> {
  const cacheKey = graphToken.substring(0, 16);
  const cached = signatureCache.get(cacheKey);
  if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
    log("Using cached signature");
    return cached.html;
  }

  // 1️⃣ Try OneDrive signature.txt
  const oneDriveSig = await fetchSignatureFromOneDrive(graphToken);
  if (oneDriveSig) {
    signatureCache.set(cacheKey, { html: oneDriveSig, fetchedAt: Date.now() });
    log(`OneDrive signature loaded (${oneDriveSig.length} chars)`);
    return oneDriveSig;
  }

  // 2️⃣ Fallback: extract from sent items
  const sentSig = await fetchSignatureFromSentItems(graphToken);
  if (sentSig) {
    signatureCache.set(cacheKey, { html: sentSig, fetchedAt: Date.now() });
    log(`Sent-items signature loaded (${sentSig.length} chars)`);
    return sentSig;
  }

  log("No signature found");
  return null;
}

/**
 * Appends the HTML signature to an email body.
 * ALWAYS produces an HTML body — plain text is converted to HTML first.
 * Idempotency: if the signature is already present, body is returned unchanged.
 */
export function appendSignature(
  body: string,
  bodyType: string,
  signature: string
): string {
  if (!signature || !body) return body;

  // Idempotency guard
  const anchor = extractIdempotencyAnchor(signature);
  if (anchor && body.includes(anchor)) {
    log("Signature already present — skipping injection");
    return body;
  }

  const isHtml = bodyType?.toLowerCase() === "html";

  // Convert plain text to HTML so the signature renders correctly
  let htmlBody = isHtml
    ? body
    : `<html><body><p style="font-family:Calibri,Arial,sans-serif;font-size:11pt;white-space:pre-wrap;">${escHtml(body)}</p></body></html>`;

  const sigBlock = `<div style="margin-top:20px;padding-top:16px;border-top:1px solid #d0d0d0;">${signature}</div>`;

  if (/<\/body>/i.test(htmlBody)) {
    return htmlBody.replace(/<\/body>/i, `${sigBlock}</body>`);
  }
  return `${htmlBody}${sigBlock}`;
}

export function invalidateSignatureCache(graphToken: string): void {
  signatureCache.delete(graphToken.substring(0, 16));
}

// ── OneDrive signature.txt ────────────────────────────────────────────────────

const ONEDRIVE_SIG_PATH = "Agent365-Bridge/signature.txt";

/**
 * Reads Agent365-Bridge/signature.txt from the user's OneDrive and
 * builds a structured HTML signature block with the ABK Solutions logo.
 *
 * File format (plain text, one item per line):
 *   Line 1: Full name
 *   Line 2: Job title
 *   Line 3+: m: / t: / a: / w: / e: prefixed fields, or free text
 */
async function fetchSignatureFromOneDrive(graphToken: string): Promise<string | null> {
  try {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_SIG_PATH}:/content`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${graphToken}` },
    });

    if (resp.status === 404) {
      log("No OneDrive signature.txt found — will try sent items");
      return null;
    }
    if (!resp.ok) {
      log(`OneDrive read failed: ${resp.status}`);
      return null;
    }

    const text = (await resp.text()).trim();
    if (!text) return null;

    return buildSignatureHtml(text);
  } catch (err) {
    log(`OneDrive signature fetch error: ${err}`);
    return null;
  }
}

// ── HTML builder ──────────────────────────────────────────────────────────────

function buildSignatureHtml(sigText: string): string {
  const lines = sigText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  const lineHtml = lines.map((line, i) => {
    if (i === 0) return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:bold;color:#1a1a1a;padding:0 0 2px 0;">${escHtml(line)}</td></tr>`;
    if (i === 1) return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:9.5pt;color:#666666;padding:0 0 6px 0;">${escHtml(line)}</td></tr>`;

    const labelMatch = line.match(/^([mtawe]):\s*(.+)$/i);
    if (labelMatch) {
      const label = labelMatch[1].toLowerCase();
      const value = labelMatch[2].trim();
      const labelStr = `<span style="color:#0066CC;font-weight:600;">${label}:</span>`;
      if (label === "e") {
        return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:9pt;padding:0 0 1px 0;">${labelStr} <a href="mailto:${escHtml(value)}" style="color:#0066CC;text-decoration:none;">${escHtml(value)}</a></td></tr>`;
      }
      if (label === "w") {
        const href = value.startsWith("http") ? value : `https://${value}`;
        return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:9pt;padding:0 0 1px 0;">${labelStr} <a href="${escHtml(href)}" style="color:#0066CC;text-decoration:none;">${escHtml(value)}</a></td></tr>`;
      }
      return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:9pt;padding:0 0 1px 0;">${labelStr} ${escHtml(value)}</td></tr>`;
    }

    // Disclaimer or free text — small italic grey
    if (line.length > 80) {
      return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:8pt;color:#888888;font-style:italic;padding-top:8px;">${escHtml(line)}</td></tr>`;
    }

    return `<tr><td style="font-family:Calibri,Arial,sans-serif;font-size:9pt;color:#444444;padding:0 0 1px 0;">${escHtml(line)}</td></tr>`;
  }).join("\n");

  const logoRow = LOGO_BASE64
    ? `<tr><td style="padding:0 0 10px 0;"><img src="cid:${LOGO_CID}" alt="ABK Solutions" width="180" height="58" style="display:block;border:0;" /></td></tr>`
    : `<tr><td style="padding:0 0 10px 0;font-family:Calibri,Arial,sans-serif;font-size:13pt;font-weight:bold;color:#0066CC;">ABK Solutions</td></tr>`;

  return `<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
  <tbody>
  ${logoRow}
  ${lineHtml}
  </tbody>
</table>`;
}

function escHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ── Sent-items fallback ───────────────────────────────────────────────────────

async function fetchSignatureFromSentItems(graphToken: string): Promise<string | null> {
  try {
    const url =
      "https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages" +
      "?$top=20&$select=id,body&$expand=attachments&$orderby=sentDateTime+desc";

    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${graphToken}`, Accept: "application/json" },
    });

    if (resp.status === 403 || resp.status === 401) {
      log(`Mail.Read not authorized (${resp.status})`);
      return null;
    }
    if (!resp.ok) {
      log(`sentItems fetch failed: ${resp.status}`);
      return null;
    }

    const data = await resp.json() as {
      value?: Array<{
        id?: string;
        body?: { contentType?: string; content?: string };
        attachments?: Array<{ contentId?: string; contentType?: string; contentBytes?: string }>;
      }>;
    };

    for (const msg of data.value ?? []) {
      const contentType = msg.body?.contentType ?? "";
      let content = msg.body?.content ?? "";
      if (contentType.toLowerCase() !== "html" || !content) continue;

      // Replace cid: references with inline base64 data URIs
      for (const att of msg.attachments ?? []) {
        if (att.contentId && att.contentBytes && att.contentType) {
          const cid = att.contentId.replace(/[<>]/g, "");
          content = content.replace(
            new RegExp(`cid:${cid}`, "g"),
            `data:${att.contentType};base64,${att.contentBytes}`
          );
        }
      }

      // Skip if cid: references are still unresolved
      if (/cid:[a-zA-Z0-9]/i.test(content)) {
        log("Skipping email with unresolved cid: references");
        continue;
      }

      const sig = extractSignature(content);
      if (sig) return sig;
    }

    return null;
  } catch (err) {
    log(`Error fetching sent items: ${err}`);
    return null;
  }
}

// ── Signature extraction from HTML email ──────────────────────────────────────

function extractSignature(html: string): string | null {
  const sigById = extractByDivId(html, "Signature");
  if (sigById) { log("Extracted via div#Signature"); return sigById; }

  const appendOnSend = extractByDivId(html, "appendonsend");
  if (appendOnSend) { log("Extracted via div#appendonsend"); return appendOnSend; }

  const afterWrapper = extractAfterMainWrapper(html);
  if (afterWrapper) { log("Extracted after divtagdefaultwrapper"); return afterWrapper; }

  const byLogo = extractByCompanyLogo(html);
  if (byLogo) { log("Extracted via ABK logo anchor"); return byLogo; }

  const afterHr = extractAfterLastHr(html);
  if (afterHr) { log("Extracted after <hr>"); return afterHr; }

  return null;
}

function extractByDivId(html: string, id: string): string | null {
  const re = new RegExp(`<div[^>]+id=["']${id}["'][^>]*>`, "i");
  const match = re.exec(html);
  if (!match) return null;
  const start = match.index;
  let depth = 1;
  let i = start + match[0].length;
  while (i < html.length && depth > 0) {
    const openIdx = html.indexOf("<div", i);
    const closeIdx = html.indexOf("</div>", i);
    if (closeIdx < 0) break;
    if (openIdx >= 0 && openIdx < closeIdx) { depth++; i = openIdx + 4; }
    else { depth--; if (depth === 0) return html.substring(start, closeIdx + 6).trim() || null; i = closeIdx + 6; }
  }
  return null;
}

function extractAfterMainWrapper(html: string): string | null {
  const wrapperRe = /<div[^>]+id=["']divtagdefaultwrapper["'][^>]*>/i;
  const wrapperMatch = wrapperRe.exec(html);
  if (!wrapperMatch) return null;
  let depth = 1;
  let i = wrapperMatch.index + wrapperMatch[0].length;
  while (i < html.length && depth > 0) {
    const openIdx = html.indexOf("<div", i);
    const closeIdx = html.indexOf("</div>", i);
    if (closeIdx < 0) break;
    if (openIdx >= 0 && openIdx < closeIdx) { depth++; i = openIdx + 4; }
    else { depth--; i = closeIdx + 6; }
  }
  const after = html.substring(i).replace(/<\/body>[\s\S]*$/i, "").trim();
  return after.replace(/<[^>]+>/g, "").trim().length > 20 ? after : null;
}

function extractByCompanyLogo(html: string): string | null {
  const imgRe = /<img[^>]+(src|alt)=["'][^"']*abk[^"']*["'][^>]*>/i;
  const imgMatch = imgRe.exec(html);
  if (!imgMatch) return null;
  const lastDivStart = html.substring(0, imgMatch.index).lastIndexOf("<div");
  if (lastDivStart < 0) return null;
  const fromDiv = html.substring(lastDivStart);
  const bodyClose = fromDiv.search(/<\/body>/i);
  return (bodyClose >= 0 ? fromDiv.substring(0, bodyClose) : fromDiv).trim() || null;
}

function extractAfterLastHr(html: string): string | null {
  const hrIdx = html.lastIndexOf("<hr");
  if (hrIdx < 0) return null;
  const after = html.substring(hrIdx).replace(/<\/body>[\s\S]*$/i, "").trim();
  return after.replace(/<[^>]+>/g, "").trim().length > 20 ? after : null;
}

function extractIdempotencyAnchor(signature: string): string | null {
  const text = signature.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
  return text.length >= 20 ? text.substring(0, 40) : null;
}
