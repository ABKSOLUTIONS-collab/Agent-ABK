import { ClientSecretCredential } from "@azure/identity";
import { getResetEmailSender } from "../auth/app-store";

// ── App-only Graph email sender ────────────────────────────────────────────
// Used for system emails (password reset) triggered by anonymous visitors,
// where no delegated user token is available. Requires the Azure AD app
// registration to be granted the Mail.Send APPLICATION permission with
// tenant admin consent — this cannot be done from code, see README.

const TENANT_ID     = process.env.AZURE_TENANT_ID ?? "";
const CLIENT_ID     = process.env.AZURE_CLIENT_ID ?? "";
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET ?? "";

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const REFRESH_BUFFER_MS = 60_000;

let credential: ClientSecretCredential | null = null;
let cachedToken: { token: string; expiresOnMs: number } | null = null;

function getCredential(): ClientSecretCredential {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error(
      "AZURE_TENANT_ID, AZURE_CLIENT_ID and AZURE_CLIENT_SECRET must be set for app-only Graph email sending"
    );
  }
  if (!credential) {
    credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  }
  return credential;
}

async function getAppOnlyGraphToken(): Promise<string> {
  if (cachedToken && Date.now() < cachedToken.expiresOnMs - REFRESH_BUFFER_MS) {
    return cachedToken.token;
  }
  const result = await getCredential().getToken(GRAPH_SCOPE);
  if (!result) {
    throw new Error("Failed to acquire app-only Graph token (Mail.Send application permission may be missing/unconsented)");
  }
  cachedToken = { token: result.token, expiresOnMs: result.expiresOnTimestamp };
  return result.token;
}

export async function sendSystemEmail(to: string, subject: string, html: string): Promise<void> {
  const sender = await getResetEmailSender();
  const token = await getAppOnlyGraphToken();

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: [{ emailAddress: { address: to } }],
        },
        saveToSentItems: false,
      }),
    }
  );

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph sendMail failed (${res.status}) from ${sender}: ${body}`);
  }
}
