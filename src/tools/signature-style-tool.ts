/**
 * signature-style-tool.ts
 *
 * MCP tool: GetUserEmailSignatureStyle
 *
 * DUAL PURPOSE — same tool, two modes:
 *
 *   MODE 1 – SAVE (called with a "name" argument)
 *     Writes the user's details to OneDrive: Agent365-Bridge/signature.txt
 *     Then returns the ready-to-use HTML signature block.
 *     Use this the first time, or when the user wants to update their signature.
 *
 *   MODE 2 – READ (called with no arguments)
 *     Reads Agent365-Bridge/signature.txt (or falls back to the user's Outlook
 *     sent items) and returns the ready-to-use HTML block, or a "not set up
 *     yet" message telling the caller to collect details and call again.
 *
 * Auto-injection also happens transparently in mcp-proxy-server.ts, but ONLY
 * from the OneDrive signature.txt (never the sent-items guess below) — so a
 * user with no saved signature gets none auto-attached, which is what forces
 * the agent to notice and go through this tool's ask-once flow instead of
 * silently relying on a possibly-wrong extraction from an old email.
 * The agent should NOT paste the returned HTML into the email body itself —
 * once saved, it's attached automatically.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getUserSignature, getUserSignatureStrict, saveUserSignature } from "./signature-service";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [sig-tool] ${msg}\n`);
}

// ── Tool definition ───────────────────────────────────────────────────────────

export const SIGNATURE_STYLE_TOOL_NAME = "GetUserEmailSignatureStyle";

export const SIGNATURE_STYLE_TOOL: Tool = {
  name: SIGNATURE_STYLE_TOOL_NAME,
  description: `Sets up or returns the user's email signature (ABK Solutions logo + personal details), stored as a plain-text file in their OneDrive (Agent365-Bridge/signature.txt). Once set up, it's injected automatically into every outgoing email — you never need to paste it into the email body yourself.

YOU MUST CALL THIS TOOL (without arguments) before the FIRST email you send, draft, reply to, or forward in a conversation — before SendEmailWithAttachments, CreateDraftMessage, ReplyToMessage, ReplyAllToMessage, ForwardMessage, ReplyWithFullThread, ReplyAllWithFullThread, or ForwardMessageWithFullThread. If it reports no signature is set up yet, ask the CURRENTLY SIGNED-IN USER (the person you are chatting with) for THEIR OWN full name, job title, and (optionally) phone, mobile, email and any extra line (address/disclaimer) — then call this tool again WITH those details to save it. After that, proceed with sending the email; the signature is attached automatically. Only ask once per user — after it's saved, subsequent calls in this or any later conversation will find it on OneDrive.

IMPORTANT: the details you save are for the SIGNED-IN USER ONLY, never for an email recipient or anyone else you've seen mentioned (e.g. in a previous email's signature or contact card). Never infer or reuse another person's name/title/contact details for this — always ask the signed-in user directly if you don't already know their own details from this conversation. Saving a mismatched email is rejected by the server.

• Called WITHOUT arguments → reads the existing signature. Returns its HTML, or a message asking you to collect the details above.
• Called WITH a "name" argument → SAVES the given details (the signed-in user's own) to OneDrive, then returns the HTML. Also use this if the user asks to update their signature.`,
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
export const SET_SIGNATURE_TOOL = SIGNATURE_STYLE_TOOL;

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

// ── Handler ───────────────────────────────────────────────────────────────────

/**
 * Dual-purpose handler.
 * If args.name is provided  → save signature.txt to OneDrive, then return HTML.
 * If no args (or no name)   → read the existing signature (OneDrive, then sent items).
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
    if (phone)  lines.push(`t: ${phone}`);
    if (mobile) lines.push(`m: ${mobile}`);
    if (emailV) lines.push(`e: ${emailV}`);
    if (extra)  lines.push(extra);

    log(`Saving signature for ${name}`);
    const result = await saveUserSignature(graphToken, lines);
    if (!result.ok) {
      return `Error saving signature to OneDrive: ${result.error}`;
    }
    return `Signature saved. It will be used automatically on future emails.\n\n${result.html}`;
  }

  // ── READ MODE ──────────────────────────────────────────────────────────────
  log("GetUserEmailSignatureStyle called (read)");

  // Saved (OneDrive) signature — this is what auto-injection will use too.
  const savedSig = await getUserSignatureStrict(graphToken);
  if (savedSig) {
    return `Your saved email signature (HTML) — already attached automatically on outgoing emails, no action needed:\n\n${savedSig}`;
  }

  // Nothing saved yet — check if one can be guessed from sent items, but
  // don't treat that as "done": still ask the user to confirm/save one.
  const guessedSig = await getUserSignature(graphToken);
  if (!guessedSig) {
    return "No signature set up yet. Ask the user for their full name, job title, and (optionally) phone, mobile, email and an extra line (address/disclaimer), then call GetUserEmailSignatureStyle again with those as arguments (name, title, phone, mobile, email, extra) to set it up.";
  }

  return `No signature saved yet — this was guessed from an old sent email and is NOT reliable and NOT auto-attached. Ask the user for their full name, job title, and (optionally) phone, mobile, email and an extra line, then call GetUserEmailSignatureStyle again with those as arguments to save a proper one. Guessed signature for reference:\n\n${guessedSig}`;
}

// Kept for proxy-server handler compatibility
export async function handleSetSignature(
  graphToken: string,
  args: Record<string, unknown>
): Promise<string> {
  return handleGetSignatureStyle(graphToken, args);
}
