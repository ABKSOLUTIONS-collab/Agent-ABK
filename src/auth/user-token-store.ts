import { TableClient, AzureNamedKeyCredential, TableEntity } from "@azure/data-tables";

// ── Azure Table Storage config ────────────────────────────────────────────────
const ACCOUNT     = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY = process.env.AZURE_STORAGE_KEY ?? "";
const TABLE       = process.env.TOKEN_TABLE_NAME || "agent365tokens";

// ── Azure AD config (for silent token refresh) ────────────────────────────────
const TENANT_ID     = process.env.AZURE_TENANT_ID ?? "";
const CLIENT_ID     = process.env.AZURE_CLIENT_ID ?? "";
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET ?? "";

// WorkIQ scope (same value as MCP_PLATFORM_AUTHENTICATION_SCOPE in the container)
const WORKIQ_SCOPE = process.env.MCP_PLATFORM_AUTHENTICATION_SCOPE
  ?? "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1/.default";

const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "https://graph.microsoft.com/Mail.ReadWrite",    // for signature extraction from sent items
  "offline_access",
];

// Refresh 5 minutes before expiry
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
}

function sleep(ms: number): Promise<void> {
  return new Promise(r => setTimeout(r, ms));
}

// ── Table Storage helpers ─────────────────────────────────────────────────────

function getClient(): TableClient {
  if (!ACCOUNT || !ACCOUNT_KEY) {
    throw new Error("AZURE_STORAGE_ACCOUNT and AZURE_STORAGE_KEY must be set");
  }
  const cred = new AzureNamedKeyCredential(ACCOUNT, ACCOUNT_KEY);
  return new TableClient(
    `https://${ACCOUNT}.table.core.windows.net`,
    TABLE,
    cred
  );
}

function toRowKey(sessionToken: string): string {
  return Buffer.from(sessionToken).toString("hex").substring(0, 256);
}

interface TokenEntity extends TableEntity {
  sessionToken: string;
  accessToken:  string;
  graphToken:   string;
  refreshToken: string;
  expiresAt:    number;
  email:        string;
}

async function getEntry(sessionToken: string): Promise<TokenEntity | null> {
  try {
    const client = getClient();
    const entity = await client.getEntity<TokenEntity>("tokens", toRowKey(sessionToken));
    return entity;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) {
      log(`Table Storage getEntry error: ${err}`);
    }
    return null;
  }
}

// ── Token refresh ─────────────────────────────────────────────────────────────

async function refreshWithScopes(
  refreshToken: string,
  scopes: string[],
  email?: string
): Promise<{ accessToken: string; expiresIn: number; refreshToken?: string } | null> {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    log("Token refresh: missing AZURE_CLIENT_SECRET / AZURE_TENANT_ID / AZURE_CLIENT_ID in env");
    return null;
  }
  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id:     CLIENT_ID,
          client_secret: CLIENT_SECRET,
          grant_type:    "refresh_token",
          refresh_token: refreshToken,
          scope:         scopes.join(" "),
        }).toString(),
      }
    );
    if (!res.ok) {
      const err = await res.text();
      log(`Token refresh failed for ${email ?? "unknown"} (scope: ${scopes[0]}): ${err}`);
      return null;
    }
    const data = await res.json() as {
      access_token:   string;
      expires_in:     number;
      refresh_token?: string;
    };
    log(`Token refreshed for ${email ?? "unknown"} (scope: ${scopes[0]})`);
    return {
      accessToken:  data.access_token,
      expiresIn:    data.expires_in,
      refreshToken: data.refresh_token,
    };
  } catch (e) {
    log(`Token refresh error: ${e}`);
    return null;
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

export async function storeUserToken(
  sessionToken: string,
  accessToken:  string,
  expiresInSeconds: number,
  refreshToken?: string,
  email?: string,
  graphToken?: string
): Promise<void> {
  const client = getClient();
  const entity: TokenEntity = {
    partitionKey: "tokens",
    rowKey:       toRowKey(sessionToken),
    sessionToken,
    accessToken,
    graphToken:   graphToken   ?? "",
    refreshToken: refreshToken ?? "",
    expiresAt:    Date.now() + expiresInSeconds * 1000,
    email:        email ?? "",
  };
  await client.upsertEntity(entity, "Replace");
  log(`Stored token for session ${sessionToken.substring(0, 8)}... (${email ?? "unknown"})${graphToken ? " [+Graph]" : ""}`);
}

// ── Lookup existing sessionToken by email ─────────────────────────────────────
// Ensures the /mcp?token=... URL never changes for a returning user.
export async function getSessionTokenByEmail(email: string): Promise<string | null> {
  try {
    const client = getClient();
    const iter = client.listEntities<TokenEntity>({
      queryOptions: {
        filter: `PartitionKey eq 'tokens' and email eq '${email.replace(/'/g, "''")}'`,
      },
    });
    for await (const entity of iter) {
      log(`Found existing session for ${email}: ${entity.sessionToken.substring(0, 8)}...`);
      return entity.sessionToken;
    }
    return null;
  } catch (e) {
    log(`getSessionTokenByEmail error: ${e}`);
    return null;
  }
}

function isExpiringSoon(expiresAt: number): boolean {
  return Date.now() > expiresAt - REFRESH_BUFFER_MS;
}

// ── getUserToken with retry ───────────────────────────────────────────────────
// Retries 3x with 500ms delay to handle the race condition where Claude.ai
// sends the first MCP initialize before Table Storage finishes writing.
// IMPORTANT: Never deletes rows — the sessionToken must survive forever
// so Claude.ai's stored connector URL always works.
export async function getUserToken(sessionToken: string): Promise<string | null> {
  const MAX_ATTEMPTS = 3;
  const RETRY_DELAY_MS = 500;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const entry = await getEntry(sessionToken);

    if (!entry) {
      if (attempt < MAX_ATTEMPTS) {
        log(`Token not found (attempt ${attempt}/${MAX_ATTEMPTS}), retrying in ${RETRY_DELAY_MS}ms...`);
        await sleep(RETRY_DELAY_MS);
        continue;
      }
      log(`Token not found after ${MAX_ATTEMPTS} attempts: ${sessionToken.substring(0, 8)}...`);
      return null;
    }

    // WorkIQ token expired — refresh with correct WorkIQ scope, also refresh Graph
    if (Date.now() > entry.expiresAt) {
      if (entry.refreshToken) {
        const [workiqRefreshed, graphRefreshed] = await Promise.all([
          refreshWithScopes(entry.refreshToken, [WORKIQ_SCOPE, "offline_access"], entry.email),
          refreshWithScopes(entry.refreshToken, GRAPH_SCOPES, entry.email),
        ]);
        if (workiqRefreshed) {
          await storeUserToken(
            sessionToken,
            workiqRefreshed.accessToken,
            workiqRefreshed.expiresIn,
            workiqRefreshed.refreshToken ?? entry.refreshToken,
            entry.email,
            graphRefreshed?.accessToken ?? entry.graphToken
          );
          log(`Expired WorkIQ token refreshed for ${entry.email}`);
          return workiqRefreshed.accessToken;
        }
      }
      // Refresh failed — row stays, return null so caller knows MS token is gone
      log(`Token expired and refresh failed for ${entry.email} — row kept, user needs re-login`);
      return null;
    }

    // Expiring soon — refresh proactively (background, non-blocking)
    if (isExpiringSoon(entry.expiresAt) && entry.refreshToken) {
      void Promise.all([
        refreshWithScopes(entry.refreshToken, [WORKIQ_SCOPE, "offline_access"], entry.email),
        refreshWithScopes(entry.refreshToken, GRAPH_SCOPES, entry.email),
      ]).then(([workiqRefreshed, graphRefreshed]) => {
        if (workiqRefreshed) {
          void storeUserToken(
            sessionToken,
            workiqRefreshed.accessToken,
            workiqRefreshed.expiresIn,
            workiqRefreshed.refreshToken ?? entry.refreshToken,
            entry.email,
            graphRefreshed?.accessToken ?? entry.graphToken
          );
        }
      });
    }

    return entry.accessToken;
  }

  return null;
}

export async function getGraphToken(sessionToken: string): Promise<string | null> {
  const entry = await getEntry(sessionToken);
  if (!entry) return null;

  if (Date.now() > entry.expiresAt) {
    if (entry.refreshToken) {
      // Refresh Graph token only — preserve the existing WorkIQ accessToken in storage.
      // getUserToken (called just before this) already handled the full refresh;
      // this path only fires if called independently after expiry.
      const graphRefreshed = await refreshWithScopes(entry.refreshToken, GRAPH_SCOPES, entry.email);
      if (graphRefreshed) {
        const remainingSecs = Math.max(0, (entry.expiresAt - Date.now()) / 1000);
        await storeUserToken(
          sessionToken,
          entry.accessToken,
          remainingSecs || graphRefreshed.expiresIn,
          graphRefreshed.refreshToken ?? entry.refreshToken,
          entry.email,
          graphRefreshed.accessToken
        );
        return graphRefreshed.accessToken;
      }
    }
    return null;
  }

  if (isExpiringSoon(entry.expiresAt) && entry.refreshToken) {
    const graphRefreshed = await refreshWithScopes(entry.refreshToken, GRAPH_SCOPES, entry.email);
    if (graphRefreshed) {
      const remainingSecs = Math.max(0, (entry.expiresAt - Date.now()) / 1000);
      await storeUserToken(
        sessionToken,
        entry.accessToken,
        remainingSecs || graphRefreshed.expiresIn,
        graphRefreshed.refreshToken ?? entry.refreshToken,
        entry.email,
        graphRefreshed.accessToken
      );
      return graphRefreshed.accessToken;
    }
  }

  return entry.graphToken || entry.accessToken;
}

export function getTokenEmail(sessionToken: string): string | undefined {
  return undefined;
}

// Async variant that actually reads the stored email from Table Storage
// (getTokenEmail above is a synchronous stub kept for backward compatibility).
export async function getStoredEmail(sessionToken: string): Promise<string | undefined> {
  const entry = await getEntry(sessionToken);
  return entry?.email || undefined;
}

export async function removeUserToken(sessionToken: string): Promise<void> {
  try {
    const client = getClient();
    await client.deleteEntity("tokens", toRowKey(sessionToken));
  } catch { /* already gone */ }
}

// pruneExpiredTokens is now a NO-OP.
// We never delete rows — session tokens must live forever so that
// Claude.ai's stored connector URLs always work, regardless of when
// the Microsoft token last expired.
export function pruneExpiredTokens(): void {
  log("pruneExpiredTokens: skipped — rows are permanent by design");
}