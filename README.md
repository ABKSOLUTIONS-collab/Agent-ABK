# Agent 365 Bridge (ABK)

A multi-tenant **HTTP MCP server** that gives Claude, ChatGPT, LibreChat, and other MCP-capable clients access to Microsoft 365 (Outlook Mail, Calendar, Word, Excel, Teams, SharePoint/OneDrive, Knowledge search, OCR) on behalf of a signed-in ABK user — plus a small org-admin console for managing who has access.

> **This document reflects the code as of 2026-07-30.** The project has evolved substantially from its original design (a local stdio bridge with device-code sign-in); it is now a hosted, multi-tenant OAuth server. If you find a mismatch between this README and the code, trust the code and update this file.

## Status

✅ Deployed and in active use at ABK as an Azure Container App, serving per-user Microsoft 365 access to Claude/ChatGPT custom connectors and to a branded internal LibreChat instance (see [LibreChat integration](#librechat-integration) below).

⚠️ Some parts of this repo are **not wired into the running server** — see [Known issues / dead code](#known-issues--dead-code) before you spend time on them.

---

## What this actually is

This is **not** a local stdio process that Claude Code launches per-session. It is a standalone **Express HTTP server** (`src/index.ts` → `McpProxyServer`, `src/proxy/mcp-proxy-server.ts`) that:

1. Acts as an **OAuth 2.0 authorization server** for MCP clients (Claude.ai/ChatGPT-style dynamic client registration + authorization-code flow), so each human user signs in with their own Microsoft 365 account once and gets a permanent personal MCP URL.
2. Serves the actual **MCP protocol** (`/mcp`) over both modern StreamableHTTP and legacy SSE transports, per-request, per-user — it is stateless, keyed by a bearer session token.
3. For most tools, calls the **Microsoft Graph API directly** with the signed-in user's delegated token (no Agent 365 / Copilot license required). For everything else, it proxies live to the **Agent 365 Tooling Gateway** using the same per-user token.
4. Hosts a small **org-admin console** (`/org-admin`) for inviting/removing ABK users and managing roles — designed to be embedded (iframed) inside the internal LibreChat deployment and sharing LibreChat's own MongoDB user collection.

```mermaid
flowchart LR
    subgraph Clients
        A[Claude.ai / ChatGPT<br/>custom connector]
        B[LibreChat<br/>(ABK-branded, separate app)]
    end
    subgraph Bridge["agent365-bridge (this repo)"]
        L["/login /authorize /callback /token<br/>(OAuth server)"]
        M["/mcp (StreamableHTTP + SSE)"]
        O["/org-admin<br/>(member management console)"]
    end
    G[Microsoft Graph API]
    P[Agent 365 Tooling Gateway]
    D[(Azure Table Storage<br/>per-user tokens)]
    N[(MongoDB<br/>LibreChat users)]

    A -- Microsoft sign-in --> L
    A -- tool calls --> M
    B -. iframes .-> O
    L --> D
    M --> D
    M -- hardcoded Graph tools --> G
    M -- everything else --> P
    O --> N
```

---

## Auth: how a user actually connects

The only **live** auth path is per-user Microsoft OAuth (delegated permissions). Everything else described in older versions of this README (device code, client credentials, on-behalf-of, bearer/mock modes) exists as leftover code from an earlier design but is **not reachable from the running server** — see [Known issues](#known-issues--dead-code).

1. User opens `GET /login` (or `/authorize`, used by MCP clients doing dynamic registration) in a browser → redirected to Microsoft sign-in.
2. `GET /callback` (`src/auth/oauth-handler.ts`) exchanges the code for:
   - a token scoped to the Agent 365 gateway (`MCP_PLATFORM_AUTHENTICATION_SCOPE`), and
   - a second Graph token (via the refresh token) with `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Mail.ReadWrite`, `offline_access`.
3. Both tokens, the refresh token, and the user's email are stored in **Azure Table Storage** (table `agent365tokens`), keyed by a permanent random `sessionToken`. Tokens are silently refreshed 5 minutes before expiry; rows are never deleted (sessions don't expire).
4. The user is handed a personal URL — `https://<host>/mcp?token=<sessionToken>` — to paste into Claude.ai / ChatGPT / another MCP client as a custom connector, or the MCP client completes the OAuth dance itself via `/register` + `/authorize` + `/token`.

Separately, `src/auth/app-store.ts` + `src/auth/app-auth.ts` implement an **unrelated** email/password admin-account system (bcrypt, JWT cookie) for the not-currently-wired `web/` admin panel — see below. It has nothing to do with the Microsoft 365 sign-in above.

---

## Tool surface: hardcoded Graph calls + live Agent 365 proxy

For history/licensing reasons, most day-to-day tools now call Microsoft Graph **directly** with the user's delegated token, bypassing the Agent 365 gateway (and its Copilot licensing requirement) entirely. Everything not in this hardcoded list still proxies live to Agent 365's MCP servers using the user's token.

| Area | File | Tools |
|---|---|---|
| Mail | `src/tools/email-graph-tools.ts` | `SearchMessages`, `GetMessage`, `CreateDraftMessage`, `UpdateDraft`, `DeleteMessage`, `UpdateMessage`, `FlagEmail`, `Reply(All)ToMessage`(+`WithFullThread`), `Forward(WithFullThread)Message`, `Get/DownloadAttachments`, `Upload(Large)Attachment`, `AddDraftAttachments`, `DeleteAttachment`, `SendDraftMessage`, `SendEmailWithAttachments` |
| Calendar | `src/tools/calendar-graph-tools.ts` | `List/ListCalendarView`, `Create/Update/DeleteEvent(ById)`, `Accept/Decline/TentativelyAcceptEvent`, `CancelEvent`, `ForwardEvent`, `FindMeetingTimes`, `GetRooms`, `GetUserDateAndTimeZoneSettings` |
| Word | `src/tools/word-graph-tools.ts` | `CreateDocument`, `GetDocumentContent_mcp_WordServer`, `AddComment`, `ReplyToComment_mcp_WordServer` |
| Excel | `src/tools/excel-graph-tools.ts` | `CreateWorkbook`, `GetDocumentContent_mcp_ExcelServer`, `CreateComment`, `ReplyToComment_mcp_ExcelServer` |
| Teams meetings | `src/tools/teams-graph-tools.ts` | `GetOnlineMeetingTranscripts`, `GetOnlineMeetingAttendanceReports`, `GetOnlineMeetingAiInsights` |
| SharePoint / OneDrive | `src/tools/sharepoint-tools.ts` | `create_sharepoint_folder`, `list_sharepoint_folder`, `upload_file_to_sharepoint`, `move_sharepoint_file`, `delete_sharepoint_file`, `create_onedrive_folder`, `upload_file_to_onedrive` |
| Knowledge | `src/tools/knowledge-graph-tools.ts` | `query/configure/delete/ingest/retrieve_federated_knowledge` |
| OCR | `src/tools/ocr-tool.ts` | `ocr_search_and_read` (Azure Document Intelligence over SharePoint/OneDrive files) |
| Signature | `src/tools/signature-style-tool.ts` | `GetUserEmailSignatureStyle` / `SetMyEmailSignature` (stored on the user's OneDrive as `Agent365-Bridge/signature.txt`) |

**Everything else** — Teams chat/channels, SharePoint Lists, PowerPoint, Dataverse, user-profile ("Me"), Copilot/Agent Directory search — comes from live discovery of the 14 servers in [`ToolingManifest.json`](ToolingManifest.json) via `src/discovery/server-discovery.ts` and `src/proxy/tool-forwarder.ts`, using the signed-in user's own Agent 365 token (25-minute per-user discovery cache).

### Email signature + inline logo

`CreateDraftMessage`, `SendDraftMessage`, and `SendEmailWithAttachments` are intercepted centrally in `mcp-proxy-server.ts` before reaching their handlers: the call is **rejected** if the user hasn't saved a signature yet, otherwise the signature HTML is appended and the ABK logo (`assets/abk-logo-email.png`) is attached inline via `contentId` so Outlook renders it correctly instead of as a regular attachment.

---

## Org-admin console (`/org-admin`)

`src/admin/org-admin.ts` is the **only admin surface actually running in production**. It is built to be embedded as an iframe inside the ABK LibreChat deployment and shares LibreChat's own MongoDB `users` collection (`MONGO_URI`). It provides:

- Its own Microsoft SSO (`/org-admin/api/sso/start` / `/callback`) — independent of the parent LibreChat session.
- `GET/POST/DELETE/PATCH /org-admin/api/users` — invite (`@abk.gr` only), remove, promote/demote (`OWNER`/`ADMIN`/`USER`), revoke a user's stored Microsoft 365 token.
- Full password-reset flow for LibreChat's own passwordless-invite accounts, emailed via app-only Graph send.
- `postMessage`s the current user's role/session back to the parent LibreChat window (`abk_org_session` / `abk_org_role`).

## LibreChat integration

LibreChat integration is real but **lives partly outside this branch**:

- In this repo (`main`): `org-admin.ts` is written directly against LibreChat's MongoDB schema and reads `LIBRECHAT_URL` to build sign-in links — it assumes it's running alongside a LibreChat instance.
- The **`librechat` branch** of this same repo is a *separate application*: an ABK-branded fork of Microsoft's official LibreChat image (`Dockerfile: FROM agent365registry.azurecr.io/librechat:latest`), patched at build time (`patch.js`) to inject ABK branding/colors and restrict registration to `abk.gr` (`librechat.yaml`). It shares no source files with `main` — it's deployed as its own container, and the bridge (`agent365-bridge`) is registered as an MCP connector *inside* that LibreChat instance, with `/org-admin` iframed into its admin UI.

If you're working across both apps, check out `origin/librechat` separately — it won't appear when browsing `main`.

---

## Setup & deployment

### Prerequisites

| Requirement | Purpose |
|---|---|
| Node.js ≥ 18 | Runtime |
| Azure AD App Registration | OAuth (delegated permissions) |
| Azure Storage Account | Per-user token storage (Table Storage) + tools-cache (Blob Storage) |
| Azure Document Intelligence resource | `ocr_search_and_read` tool |
| MongoDB (LibreChat's DB) | `/org-admin` console only |
| Docker + Azure CLI (`az`) | Building/deploying via `deploy.sh` |

### Azure AD App Registration

1. **Entra ID → App registrations → New registration**, single-tenant, no redirect URI needed at creation (add `https://<host>/callback` and `https://<host>/org-admin/api/sso/callback` under **Authentication → Web** afterwards).
2. **Certificates & secrets** → new client secret → `AZURE_CLIENT_SECRET`.
3. **API permissions** → add delegated Microsoft Graph scopes: `Mail.ReadWrite`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `offline_access`, `openid`, `profile`, `email` — plus the Agent 365 gateway scope (`McpServers.*.All`, see the table in the original setup notes below) for whichever Agent 365 servers you want proxied.
4. **Authentication** → enable the redirect URIs above; public client flows are **not** required (this is a confidential client using authorization code + client secret, not device code).
5. Grant admin consent.

### Environment variables

| Variable | Required for | Notes |
|---|---|---|
| `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` | Core OAuth | From the app registration |
| `AZURE_STORAGE_ACCOUNT`, `AZURE_STORAGE_KEY` | Per-user tokens, tools cache, app-store tables | **Without these, `/callback` throws the moment a user tries to sign in.** |
| `TOKEN_TABLE_NAME` | optional | Defaults to `agent365tokens` |
| `MCP_PLATFORM_ENDPOINT`, `MCP_PLATFORM_AUTHENTICATION_SCOPE` | Agent 365 gateway proxying | Defaults point at the production gateway |
| `AGENTIC_APP_ID` | optional | Only for gateway-based server discovery |
| `SERVER_BASE_URL` | OAuth redirect construction | Must match the deployed public URL |
| `AZURE_DI_ENDPOINT`, `AZURE_DI_KEY` | `ocr_search_and_read` tool | Azure Document Intelligence resource |
| `MONGO_URI` | `/org-admin` console | LibreChat's MongoDB connection string |
| `LIBRECHAT_URL` | `/org-admin` invite emails | Public URL of the LibreChat instance |
| `PRIMARY_OWNER_EMAIL` | `/org-admin` | Email pinned as the un-demotable `OWNER` |
| `ORG_SSO_CLIENT_ID`, `ORG_SSO_CLIENT_SECRET` | `/org-admin` SSO | Can reuse the same app registration |
| `ORG_ADMIN_ALLOWED_ORIGINS` | `/org-admin` CORS | Comma-separated origins allowed to iframe/call it |
| `JWT_SECRET` | `web/` admin panel session cookies | Only relevant if you re-wire `admin-routes.ts`/`auth-routes.ts` (currently dead — see below) |
| `PORT` | optional | Defaults to `3000` |
| `NODE_ENV` | | `production` in deployment |

`AUTH_MODE`, `BEARER_TOKEN`, and `MCP_API_KEY` are read/required by legacy scripts or `deploy.sh` but **not consumed by the running server** — see [Known issues](#known-issues--dead-code).

### Run locally

```bash
npm install
npm run build
npm run dev      # or: npm start (after build)
```

The server listens on `http://localhost:3000`; visit `/login` to test the OAuth flow, `/health` for a status check.

### Deploy (Azure Container Apps)

`deploy.sh` is the real, current deployment path: builds and pushes the image to ACR (`Agent365Registry`), provisions/reuses the `abkagent365storage` Storage Account, and creates/updates the `agent365-bridge` Container App (resource group `ABKAgent365`) with HTTP ingress on port 3000. It sets the core OAuth + storage + Document Intelligence env vars — **it does not set `MONGO_URI`, `JWT_SECRET`, `ORG_SSO_*`, `LIBRECHAT_URL`, or `PRIMARY_OWNER_EMAIL`**, so those must be maintained separately (Container App secrets/portal) if `/org-admin` is to keep working after a redeploy.

```bash
export AZURE_TENANT_ID=... AZURE_CLIENT_ID=... AZURE_CLIENT_SECRET=...
export MCP_API_KEY=... AZURE_DI_ENDPOINT=... AZURE_DI_KEY=...
./deploy.sh
```

`Dockerfile` builds only the Node backend (`dist/`) — the `web/` React app is not part of the image.

---

## npm scripts

| Script | Status | Notes |
|---|---|---|
| `npm run build` | ✅ live | `tsc` |
| `npm run dev` | ✅ live | `ts-node src/index.ts` — runs the real server |
| `npm run start` | ✅ live | `node dist/index.js` — what the container runs |
| `npm run clean` | ✅ live | `rimraf dist` |
| `npm run setup` | ⚠️ stale | Interactive wizard for the old A365-CLI/device-code flow; not applicable to the OAuth server |
| `npm run login` / `npm run logout` | ⚠️ stale | Old device-code login writing to `~/.agent365-bridge/`; unrelated to the live per-user Table Storage tokens |
| `npm run register` | ❌ broken for this architecture | Registers the bridge with Claude Code CLI as a **stdio** server (`claude mcp add --transport stdio ...`), but `dist/index.js` is now an HTTP server that never speaks stdio JSON-RPC. Don't use this — add the bridge as a custom HTTP/OAuth connector instead. |
| `npm run mock` | ⚠️ stale | Built for the old manifest-discovery flow |
| `npm run test:e2e` | ❌ broken | Spawns `dist/index.js` over stdio and expects an MCP handshake; fails immediately against the current HTTP server |

`web/` (`npm run dev` / `build` / `lint` / `preview`) all work in isolation, but nothing in the deployed server currently serves this frontend — see below.

---

## Known issues / dead code

- **`web/` React admin app is fully built but not wired in.** Its backend, `src/routes/admin-routes.ts` and `src/routes/auth-routes.ts`, are never imported/mounted by `mcp-proxy-server.ts` or `index.ts`. The `/org-admin` console (MongoDB-backed) appears to have superseded it. Either wire the routes back in or remove the dead code — don't assume the admin UI works against the current server as-is.
- **`register-claude.ts` and `test:e2e` assume a stdio server** that no longer exists; both will misbehave (see npm scripts table).
- **`src/auth/token-provider.ts`, `token-cache.ts`, `auth-record-cache.ts`** (device code / client credentials / bearer / mock auth) are only referenced by the legacy `scripts/*.ts` files, not by the live server. `AUTH_MODE` is set in deployment configs but nothing in the live request path reads it.
- **No OBO (on-behalf-of) implementation exists** despite being documented in earlier versions of this README.
- **`mcp-proxy-server.ts` logs "Loaded N tools from Table Storage cache"** — the tools cache is actually Azure **Blob** Storage; only per-user OAuth tokens live in Table Storage.
- **`MCP_API_KEY`** is required by `deploy.sh` but not read anywhere in `src/` — confirm whether it's still needed before assuming it does something.
- **`containerapp.yaml`** is a generic template with placeholder values and comments claiming "MCP is stdio, no ingress needed" — both are stale; the real, current export is `app.yaml`, and ingress is genuinely external on port 3000.

---

## Disclaimer

**Authentication & Liability**: This project uses ABK's own Azure AD App Registration and operates under the context of the signed-in user. Whoever operates this deployment is responsible for the security of client secrets, storage keys, and tokens. No warranty is made regarding data loss, security breaches, or unexpected charges arising from its use.
