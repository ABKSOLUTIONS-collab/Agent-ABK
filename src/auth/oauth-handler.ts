import * as crypto from "crypto";
import { Express, Request, Response } from "express";
import { storeUserToken, getSessionTokenByEmail } from "./user-token-store";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
}

interface OAuthConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  serverBaseUrl: string;
  mcpScope: string;
}

interface PendingAuth {
  state: string;
  codeVerifier: string;
  sessionToken: string;
  createdAt: number;
  redirectUri?: string;
}

const pendingAuths = new Map<string, PendingAuth>();
const registeredClients = new Map<string, { clientId: string; clientSecret: string; redirectUris: string[] }>();

function generateToken(bytes = 32): string {
  return crypto.randomBytes(bytes).toString("hex");
}

async function getGraphTokenFromRefresh(
  tenantId: string,
  clientId: string,
  clientSecret: string,
  refreshToken: string
): Promise<string | null> {
  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: "refresh_token",
          refresh_token: refreshToken,
          scope: [
            "https://graph.microsoft.com/Files.ReadWrite.All",
            "https://graph.microsoft.com/Sites.ReadWrite.All",
            "https://graph.microsoft.com/Mail.ReadWrite",  // for signature extraction
            "offline_access",
          ].join(" "),
        }).toString(),
      }
    );
    if (!res.ok) {
      const err = await res.text();
      log(`Graph token from refresh failed: ${err}`);
      return null;
    }
    const data = await res.json() as { access_token: string };
    log("Graph token acquired via refresh token");
    return data.access_token;
  } catch (e) {
    log(`Graph token refresh error: ${e}`);
    return null;
  }
}

export function registerOAuthEndpoints(app: Express, config: OAuthConfig): void {
  const { tenantId, clientId, clientSecret, serverBaseUrl, mcpScope } = config;
  const EXPOSED_SCOPE = "mcp:access";

  // Session tokens never expire from Claude.ai's perspective — they are
  // permanent identifiers backed by Table Storage. The actual Microsoft
  // tokens are refreshed silently inside getUserToken/getGraphToken.
  const SESSION_EXPIRES_IN = 3600 * 24 * 365 * 10; // 10 years

  const WORKIQ_SCOPES = [mcpScope, "offline_access", "openid", "profile", "email"];

  app.get("/.well-known/oauth-authorization-server", (_req, res) => {
    log("OAuth: metadata discovery");
    res.json({
      issuer: serverBaseUrl,
      authorization_endpoint: `${serverBaseUrl}/authorize`,
      token_endpoint: `${serverBaseUrl}/token`,
      registration_endpoint: `${serverBaseUrl}/register`,
      scopes_supported: [EXPOSED_SCOPE],
      response_types_supported: ["code"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      code_challenge_methods_supported: ["S256"],
      token_endpoint_auth_methods_supported: ["client_secret_post", "none"],
    });
  });

  app.post("/register", (req, res) => {
    const { redirect_uris, client_name } = req.body ?? {};
    if (!redirect_uris?.length) {
      res.status(400).json({ error: "redirect_uris required" });
      return;
    }
    const newClientId = `claude-${generateToken(8)}`;
    const newClientSecret = generateToken(16);
    registeredClients.set(newClientId, {
      clientId: newClientId,
      clientSecret: newClientSecret,
      redirectUris: redirect_uris,
    });
    log(`OAuth: registered client ${newClientId} (${client_name ?? "unknown"})`);
    res.status(201).json({
      client_id: newClientId,
      client_secret: newClientSecret,
      redirect_uris,
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      token_endpoint_auth_method: "client_secret_post",
    });
  });

  app.get("/authorize", (req: Request, res: Response) => {
    const { client_id, redirect_uri, state, code_challenge } = req.query as Record<string, string>;
    if (!client_id || !redirect_uri || !state) {
      res.status(400).send("Missing required parameters");
      return;
    }
    const sessionToken = generateToken(32);
    const ourState = generateToken(16);
    pendingAuths.set(ourState, {
      state,
      codeVerifier: code_challenge ?? "",
      sessionToken,
      createdAt: Date.now(),
      redirectUri: redirect_uri,
    });
    const cutoff = Date.now() - 10 * 60 * 1000;
    for (const [key, val] of pendingAuths.entries()) {
      if (val.createdAt < cutoff) pendingAuths.delete(key);
    }
    const callbackUrl = `${serverBaseUrl}/callback`;
    const msAuthUrl = new URL(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`
    );
    msAuthUrl.searchParams.set("client_id", clientId);
    msAuthUrl.searchParams.set("response_type", "code");
    msAuthUrl.searchParams.set("redirect_uri", callbackUrl);
    msAuthUrl.searchParams.set("scope", WORKIQ_SCOPES.join(" "));
    msAuthUrl.searchParams.set("state", ourState);
    msAuthUrl.searchParams.set("prompt", "select_account");
    log(`OAuth: redirecting to Microsoft login (state=${ourState.substring(0, 8)}...)`);
    res.redirect(msAuthUrl.toString());
  });

  app.get("/callback", async (req: Request, res: Response) => {
    const { code, state: ourState, error, error_description } = req.query as Record<string, string>;
    if (error) {
      log(`OAuth: callback error: ${error} - ${error_description}`);
      res.status(400).send(`Authentication failed: ${error_description ?? error}`);
      return;
    }
    if (!code || !ourState) {
      res.status(400).send("Missing code or state");
      return;
    }
    const pending = pendingAuths.get(ourState);
    if (!pending) {
      res.status(400).send("Invalid or expired state");
      return;
    }
    pendingAuths.delete(ourState);
    const isDirectLogin = pending.state === "direct";

    try {
      const callbackUrl = `${serverBaseUrl}/callback`;

      // ── Step 1: Get Work IQ token ─────────────────────────────────────────
      const tokenRes = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            grant_type: "authorization_code",
            code,
            redirect_uri: callbackUrl,
            scope: WORKIQ_SCOPES.join(" "),
          }).toString(),
        }
      );

      if (!tokenRes.ok) {
        const errText = await tokenRes.text();
        log(`OAuth: token exchange failed: ${errText}`);
        res.status(500).send("Token exchange failed");
        return;
      }

      const tokenData = await tokenRes.json() as {
        access_token: string;
        refresh_token?: string;
        expires_in: number;
        id_token?: string;
      };

      // Extract email
      let email: string | undefined;
      if (tokenData.id_token) {
        try {
          const payload = JSON.parse(
            Buffer.from(tokenData.id_token.split(".")[1], "base64").toString()
          );
          email = payload.preferred_username ?? payload.email ?? payload.upn;
        } catch { /* ignore */ }
      }

      // ── Reuse existing sessionToken if this email has logged in before ────
      // This ensures the /mcp?token=... URL never changes for a returning user.
      let finalSessionToken = pending.sessionToken;
      if (email) {
        const existingToken = await getSessionTokenByEmail(email);
        if (existingToken) {
          finalSessionToken = existingToken;
          log(`OAuth: reusing existing session token for ${email}`);
        } else {
          log(`OAuth: new user ${email}, assigning token ${finalSessionToken.substring(0, 8)}...`);
        }
      }

      // ── Step 2: Use refresh token to get Graph token ──────────────────────
      let graphToken: string | undefined;
      if (tokenData.refresh_token) {
        const gt = await getGraphTokenFromRefresh(
          tenantId, clientId, clientSecret, tokenData.refresh_token
        );
        graphToken = gt ?? undefined;
      }

      // ── Step 3: Store both tokens ─────────────────────────────────────────
      await storeUserToken(
        finalSessionToken,
        tokenData.access_token,
        tokenData.expires_in,
        tokenData.refresh_token,
        email,
        graphToken
      );

      log(`OAuth: user authenticated (${email ?? "unknown"}) WorkIQ✓${graphToken ? " Graph✓" : " Graph✗"}`);

      // ── Personal URL page ─────────────────────────────────────────────────
      if (isDirectLogin) {
        const personalUrl = `${serverBaseUrl}/mcp?token=${finalSessionToken}`;
        res.send(`<!DOCTYPE html>
<html>
<head>
  <title>Agent365-Bridge — Connected</title>
  <meta charset="utf-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
           max-width: 600px; margin: 60px auto; padding: 20px; color: #1a1a1a; }
    h1 { color: #0078d4; }
    .url-box { background: #f3f3f3; border: 1px solid #ddd; border-radius: 8px;
               padding: 16px; font-family: monospace; font-size: 13px;
               word-break: break-all; margin: 20px 0; }
    .copy-btn { background: #0078d4; color: white; border: none; padding: 10px 20px;
                border-radius: 6px; cursor: pointer; font-size: 14px; }
    .copy-btn:hover { background: #005a9e; }
    .steps { background: #e8f4fd; border-radius: 8px; padding: 16px; margin-top: 20px; }
    .steps ol { margin: 8px 0; padding-left: 20px; }
    .steps li { margin: 8px 0; }
    .success { color: #107c10; font-weight: bold; }
    .badge { background: #107c10; color: white; font-size: 11px;
             padding: 2px 8px; border-radius: 4px; margin-left: 4px; }
    .badge-warn { background: #d83b01; color: white; font-size: 11px;
                  padding: 2px 8px; border-radius: 4px; margin-left: 4px; }
  </style>
</head>
<body>
  <h1>✅ Successfully Connected!</h1>
  <p class="success">
    Signed in as: ${email ?? "your Microsoft account"}
    <span class="badge">Mail ✓</span>
    <span class="badge">Calendar ✓</span>
    ${graphToken ? '<span class="badge">SharePoint ✓</span>' : '<span class="badge-warn">SharePoint ✗</span>'}
  </p>
  <p>Your personal connector URL:</p>
  <div class="url-box" id="url">${personalUrl}</div>
  <button class="copy-btn" onclick="copyUrl()">📋 Copy URL</button>
  <div class="steps">
    <strong>How to add this to Claude:</strong>
    <ol>
      <li>Go to <strong>claude.ai</strong> → Settings → Connectors</li>
      <li>Click <strong>Add custom connector</strong></li>
      <li>Name: <code>Agent365-Bridge</code></li>
      <li>URL: paste your personal URL above</li>
      <li>Click <strong>Add</strong> — done! ✅</li>
    </ol>
    <p><strong>⚠️ Keep this URL private</strong> — it gives access to your Microsoft 365 data.</p>
    <p>ℹ️ This URL is <strong>permanent</strong> — it will never change even if you sign in again.</p>
  </div>
  <script>
    function copyUrl() {
      navigator.clipboard.writeText('${personalUrl}');
      document.querySelector('.copy-btn').textContent = '✅ Copied!';
      setTimeout(() => document.querySelector('.copy-btn').textContent = '📋 Copy URL', 2000);
    }
  </script>
</body>
</html>`);
        return;
      }

      // ── Redirect back to the OAuth client (Claude, ChatGPT, Cursor, etc.) ──
      const authCode = `${finalSessionToken}:${generateToken(8)}`;
      pendingAuths.set(authCode, {
        state: pending.state,
        codeVerifier: "",
        sessionToken: finalSessionToken,
        createdAt: Date.now(),
      });
      const clientCallback = pending.redirectUri ?? `https://claude.ai/api/mcp/auth_callback`;
      const redirectUrl = new URL(clientCallback);
      redirectUrl.searchParams.set("code", authCode);
      redirectUrl.searchParams.set("state", pending.state);
      log(`OAuth: redirecting back to client: ${clientCallback}`);
      res.redirect(redirectUrl.toString());

    } catch (e) {
      log(`OAuth: callback error: ${e}`);
      res.status(500).send("Internal error during authentication");
    }
  });

  // ── Token endpoint ────────────────────────────────────────────────────────
  // Handles both authorization_code (first login) and refresh_token (renewal).
  // The sessionToken IS the refresh_token — Claude.ai can always renew itself.
  app.post("/token", async (req: Request, res: Response) => {
    const { grant_type, code, refresh_token } = req.body ?? {};

    // refresh_token grant — Claude.ai renews its session automatically
    if (grant_type === "refresh_token" && refresh_token) {
      log(`OAuth: refresh_token grant for session ${String(refresh_token).substring(0, 8)}...`);
      res.json({
        access_token: refresh_token,   // sessionToken stays the same
        refresh_token: refresh_token,  // and so does the refresh token
        token_type: "Bearer",
        expires_in: SESSION_EXPIRES_IN,
        scope: EXPOSED_SCOPE,
      });
      return;
    }

    // authorization_code grant — first login
    if (grant_type !== "authorization_code") {
      res.status(400).json({ error: "unsupported_grant_type" });
      return;
    }
    if (!code) {
      res.status(400).json({ error: "missing code" });
      return;
    }
    const pending = pendingAuths.get(code);
    if (!pending) {
      res.status(400).json({ error: "invalid_grant", error_description: "Code expired or invalid" });
      return;
    }
    pendingAuths.delete(code);
    const sessionToken = pending.sessionToken;
    log(`OAuth: token exchange for session ${sessionToken.substring(0, 8)}...`);
    res.json({
      access_token: sessionToken,
      refresh_token: sessionToken,  // same value — allows Claude.ai to auto-renew
      token_type: "Bearer",
      expires_in: SESSION_EXPIRES_IN,
      scope: EXPOSED_SCOPE,
    });
  });

  app.get("/login", (_req: Request, res: Response) => {
    const sessionToken = generateToken(32);
    const ourStateKey = generateToken(16);
    pendingAuths.set(ourStateKey, {
      state: "direct",
      codeVerifier: "",
      sessionToken,
      createdAt: Date.now(),
    });
    const callbackUrl = `${serverBaseUrl}/callback`;
    const msAuthUrl = new URL(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`
    );
    msAuthUrl.searchParams.set("client_id", clientId);
    msAuthUrl.searchParams.set("response_type", "code");
    msAuthUrl.searchParams.set("redirect_uri", callbackUrl);
    msAuthUrl.searchParams.set("scope", WORKIQ_SCOPES.join(" "));
    msAuthUrl.searchParams.set("state", ourStateKey);
    msAuthUrl.searchParams.set("prompt", "select_account");
    log("OAuth: direct login initiated (WorkIQ + Graph)");
    res.redirect(msAuthUrl.toString());
  });

  log("OAuth endpoints registered");
}