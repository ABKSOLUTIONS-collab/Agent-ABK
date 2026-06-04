/**
 * signature-style-tool.ts
 *
 * MCP tool: GetUserEmailSignatureStyle
 *
 * Reads the user's Outlook signature from their sent items (via signature-service.ts)
 * and returns the ready-to-use HTML block.
 *
 * NOTE: Auto-injection is also handled transparently in mcp-proxy-server.ts —
 * the signature is appended automatically to every outgoing email without
 * the agent needing to call this tool explicitly.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getUserSignature } from "./signature-service";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] [sig-tool] ${msg}\n`);
}

// ── Tool definition ───────────────────────────────────────────────────────────

export const SIGNATURE_STYLE_TOOL_NAME = "GetUserEmailSignatureStyle";

export const SIGNATURE_STYLE_TOOL: Tool = {
  name: SIGNATURE_STYLE_TOOL_NAME,
  description: `Returns the user's Outlook email signature as an HTML block, extracted from their sent items.
Call this tool before composing emails if you need to preview or reference the signature.
The signature is also injected automatically into every outgoing email.`,
  inputSchema: {
    type: "object" as const,
    properties: {},
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

export async function handleGetSignatureStyle(
  graphToken: string,
  _args?: Record<string, unknown>
): Promise<string> {
  log("GetUserEmailSignatureStyle called");

  const sig = await getUserSignature(graphToken);

  if (!sig) {
    return "No signature found in your sent items. Send at least one email with your signature from Outlook first.";
  }

  return `Your email signature (HTML):\n\n${sig}`;
}

// Kept for proxy-server handler compatibility
export async function handleSetSignature(
  graphToken: string,
  args: Record<string, unknown>
): Promise<string> {
  return handleGetSignatureStyle(graphToken, args);
}
