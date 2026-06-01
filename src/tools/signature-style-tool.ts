/**
 * signature-style-tool.ts
 *
 * Single MCP tool: GetUserEmailSignatureStyle
 *
 * DUAL PURPOSE — same tool, two modes:
 *
 *   MODE 1 – SAVE (called with name/title/etc.)
 *     Writes the user's details to OneDrive: Agent365-Bridge/signature.txt
 *     Then returns the ready-to-use HTML signature block.
 *     Use this the first time, or when the user wants to update their signature.
 *
 *   MODE 2 – READ (called with no arguments)
 *     Reads Agent365-Bridge/signature.txt from OneDrive.
 *     Returns the ready-to-use HTML signature block.
 *     MUST be called before every email composition tool.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getCompanyLogoBase64 } from "../auth/signature-store";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [sig-tool] ${msg}\n`);
}

const ONEDRIVE_SIG_PATH = "Agent365-Bridge/signature.txt";

// ── Tool definition ───────────────────────────────────────────────────────────

export const SIGNATURE_STYLE_TOOL_NAME = "GetUserEmailSignatureStyle";

export const SIGNATURE_STYLE_TOOL: Tool = {
  name: SIGNATURE_STYLE_TOOL_NAME,
  description: `Manages and returns the user's personal email signature (ABK Solutions logo + personal details) stored as a plain-text file in their OneDrive (Agent365-Bridge/signature.txt).

DUAL PURPOSE:
• Called WITH a "name" argument → SAVES the signature details to OneDrive, then returns the HTML block. Use this the first time or to update the signature.
• Called WITHOUT arguments → READS the saved signature from OneDrive and returns the HTML block.

YOU MUST CALL THIS TOOL (without arguments) before using SendEmailWithAttachments, CreateDraftMessage, ReplyToMessage, ReplyAllToMessage, ForwardMessage, ReplyWithFullThread, ReplyAllWithFullThread, or ForwardMessageWithFullThread. Append the returned HTML verbatim to the bottom of every email body.`,
  inputSchema: {
    type: "object" as const,
    properties: {
      name:   { type: "string", description: "Full name — provide this to SAVE or UPDATE the signature" },
      title:  { type: "string", description: "Job title / role" },
      phone:  { type: "string", description: "Phone number (optional)" },
      mobile: { type: "string", description: "Mobile number (optional)" },
      email:  { type: "string", description: "Email address (optional)" },
      extra:  { type: "string", description: "Additional lines: address, website, disclaimer, etc. (optional)" },
    },
    required: [],
  },
};

// Keep SetMyEmailSignature as a no-op alias for backward compatibility
export const SET_SIGNATURE_TOOL_NAME = "SetMyEmailSignature";
export const SET_SIGNATURE_TOOL = SIGNATURE_STYLE_TOOL; // same tool, not published separately

// Names of email composition tools whose descriptions should be patched
export const EMAIL_TOOLS_REQUIRING_SIGNATURE = new Set([
  "SendEmailWithAttachments",
  "CreateDraftMessage",
  "ReplyToMessage",
  "ReplyAllToMessage",
  "ForwardMessage",
  "ReplyWithFullThread",
  "ReplyAllWithFullThread",
  "ForwardMessageWithFullThread",
]);

// ── Main handler ──────────────────────────────────────────────────────────────

/**
 * Dual-purpose handler.
 * If args.name is provided  → save signature.txt to OneDrive, then return HTML.
 * If no args (or no name)   → read signature.txt from OneDrive and return HTML.
 */
export async function handleGetSignatureStyle(
  graphToken: string,
  args?: Record<string, unknown>
): Promise<string> {

  // ── SAVE MODE ──────────────────────────────────────────────────────────────
  if (args && typeof args.name === "string" && args.name.trim()) {
    const name   = args.name.trim();
    const title  = String(args.title  ?? "").trim();
    const phone  = String(args.phone  ?? "").trim();
    const mobile = String(args.mobile ?? "").trim();
    const emailV = String(args.email  ?? "").trim();
    const extra  = String(args.extra  ?? "").trim();

    if (!title) return "Error: 'title' is required when saving a signature.";

    const lines: string[] = [name, title];
    if (phone)  lines.push(`T: ${phone}`);
    if (mobile) lines.push(`M: ${mobile}`);
    if (emailV) lines.push(emailV);
    if (extra)  lines.push(extra);

    const fileContent = lines.join("\n");

    try {
      const uploadUrl =
        `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_SIG_PATH}:/content`;

      const resp = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${graphToken}`,
          "Content-Type": "text/plain",
        },
        body: fileContent,
      });

      if (!resp.ok) {
        const err = await resp.text();
        log(`OneDrive upload failed: ${resp.status} ${err}`);
        return `Error saving signature to OneDrive: ${resp.status} — ${err}`;
      }

      log(`signature.txt written to OneDrive for ${name}`);
    } catch (err) {
      log(`Save error: ${err}`);
      return `Error saving signature: ${err}`;
    }

    // After saving, fall through to read & return the HTML
  }

  // ── READ MODE (also runs after SAVE) ──────────────────────────────────────
  let sigText: string | null = null;
  try {
    const readUrl =
      `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_SIG_PATH}:/content`;

    const resp = await fetch(readUrl, {
      headers: { Authorization: `Bearer ${graphToken}` },
    });

    if (resp.status === 404) {
      return `No signature file found. Please call this tool again with your name, title, phone, email and address to set up your signature. Example: GetUserEmailSignatureStyle(name: "Your Name", title: "Your Title", phone: "+30...", email: "you@abk.gr", extra: "address | disclaimer text")`;
    }
    if (!resp.ok) {
      log(`OneDrive read failed: ${resp.status}`);
      return `Could not read signature from OneDrive (${resp.status}). Compose the email without a signature.`;
    }

    sigText = await resp.text();
    log(`Read signature.txt (${sigText.length} chars)`);
  } catch (err) {
    log(`Read error: ${err}`);
    return `Error reading signature: ${err}. Compose the email without a signature.`;
  }

  if (!sigText?.trim()) {
    return `Signature file exists but is empty. Call this tool with your name and title to fill it in.`;
  }

  const logoBase64 = await getCompanyLogoBase64();
  const html = buildSignatureHtml(sigText.trim(), logoBase64);

  return `Append this HTML signature verbatim to the bottom of the email body:\n\n${html}`;
}

// ── Kept for proxy-server handler compatibility ───────────────────────────────
export async function handleSetSignature(
  graphToken: string,
  args: Record<string, unknown>
): Promise<string> {
  return handleGetSignatureStyle(graphToken, args);
}

// ── HTML builder ──────────────────────────────────────────────────────────────

function buildSignatureHtml(sigText: string, logoBase64: string | null): string {
  const logoImg = logoBase64
    ? `<img src="data:image/png;base64,${logoBase64}" alt="ABK Solutions" style="height:45px;width:auto;display:block;margin-bottom:8px;"/>`
    : `<strong style="color:#0066CC;">ABK Solutions</strong><br/>`;

  const lines = sigText
    .split(/\r?\n/)
    .map((l, i) => {
      const s = l.trim();
      if (!s) return "";
      if (i === 0) return `<strong style="font-size:11pt;">${escHtml(s)}</strong>`;
      if (i === 1) return `<span style="color:#555555;font-size:9.5pt;">${escHtml(s)}</span>`;
      if (/^[TtMm]:\s/.test(s)) return `<span style="font-size:9pt;">${escHtml(s)}</span>`;
      if (/@/.test(s)) return `<span style="font-size:9pt;"><a href="mailto:${escHtml(s)}" style="color:#0066CC;text-decoration:none;">${escHtml(s)}</a></span>`;
      return `<span style="font-size:9pt;">${escHtml(s)}</span>`;
    })
    .filter(Boolean)
    .join("<br/>\n  ");

  return `<div style="margin-top:20px;padding-top:12px;border-top:2px solid #0066CC;font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#1a1a1a;line-height:1.5;">
  ${logoImg}
  ${lines}
</div>`;
}

function escHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
