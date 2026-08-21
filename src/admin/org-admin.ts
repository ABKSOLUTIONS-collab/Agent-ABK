import * as crypto from "crypto";
import { Express, Request, Response } from "express";
import { MongoClient, Db, ObjectId } from "mongodb";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { TableClient, AzureNamedKeyCredential } from "@azure/data-tables";
import { getSessionTokenByEmail, removeUserToken, getGraphToken } from "../auth/user-token-store";

const SERVER_BASE_URL = (process.env.SERVER_BASE_URL ?? "").replace(/\/$/, "");
const LIBRECHAT_URL = (process.env.LIBRECHAT_URL ?? "").replace(/\/$/, "");

// org-admin's OWN Microsoft SSO login (separate from LibreChat's — org-admin
// is served cross-origin from agent365-bridge, so it can't read LibreChat's
// session cookie; it needs its own real, verifiable identity proof instead
// of trusting anything the parent page claims). Reuses the same Entra ID
// app registration LibreChat's own SSO uses, with a second redirect URI
// registered for this callback.
const ORG_SSO_CLIENT_ID = process.env.ORG_SSO_CLIENT_ID ?? "";
const ORG_SSO_CLIENT_SECRET = process.env.ORG_SSO_CLIENT_SECRET ?? "";
const ORG_SSO_TENANT_ID = process.env.AZURE_TENANT_ID ?? "";
const ORG_SSO_REDIRECT_URI = `${SERVER_BASE_URL}/org-admin/api/sso/callback`;

function log(msg: string): void {
  process.stderr.write(`[org-admin] ${msg}\n`);
}

const PRIMARY_OWNER_EMAIL = (process.env.PRIMARY_OWNER_EMAIL ?? "").toLowerCase().trim();

type Tier = "OWNER" | "ADMIN" | "USER";

function getTier(email: string, dbRole?: string, orgRole?: string): Tier {
  if (PRIMARY_OWNER_EMAIL && email.toLowerCase() === PRIMARY_OWNER_EMAIL) return "OWNER";
  if (orgRole === "OWNER") return "OWNER";
  if (dbRole === "ADMIN" || orgRole === "ADMIN") return "ADMIN";
  return "USER";
}

interface AuthInfo {
  email: string;
  tier: Tier;
}

interface LcUser {
  _id?: ObjectId;
  email: string;
  name?: string;
  username?: string;
  role?: string;
  orgRole?: string;
  password?: string;
  createdAt?: Date;
}

// ── MongoDB ──────────────────────────────────────────────────────────────────

const MONGO_URI = process.env.MONGO_URI ?? "";
let _mongoClient: MongoClient | null = null;
let _mongoDb: Db | null = null;

async function getDb(): Promise<Db | null> {
  if (!MONGO_URI) {
    log("MONGO_URI not set");
    return null;
  }
  try {
    if (!_mongoClient) {
      _mongoClient = new MongoClient(MONGO_URI, { serverSelectionTimeoutMS: 5000 });
      await _mongoClient.connect();
      _mongoDb = _mongoClient.db();
      log("MongoDB connected");
    }
    return _mongoDb;
  } catch (e) {
    log(`MongoDB connection error: ${e}`);
    _mongoClient = null;
    _mongoDb = null;
    return null;
  }
}

async function getLcUsers(): Promise<LcUser[]> {
  const db = await getDb();
  if (!db) return [];
  try {
    return await db.collection<LcUser>("users").find({}, {
      projection: { email: 1, name: 1, username: 1, role: 1, orgRole: 1, createdAt: 1 },
    }).toArray();
  } catch (e) {
    log(`getLcUsers error: ${e}`);
    return [];
  }
}

async function getLcUserByEmail(email: string): Promise<LcUser | null> {
  const db = await getDb();
  if (!db) return null;
  try {
    return (await db.collection<LcUser>("users").findOne({ email })) ?? null;
  } catch {
    return null;
  }
}

// Real token counts each user has consumed, all-time (native LibreChat
// Transactions collection — logged for every real LLM call regardless of
// whether the balance/quota feature is enabled or not; see librechat.yaml).
// Uses rawAmount (actual prompt+completion token counts), NOT tokenValue
// (which is rawAmount scaled by a per-model pricing multiplier — the same
// inflated "credit" unit the balance feature used, confusing to show
// directly as "tokens used"). rawAmount is stored negative (a deduction),
// so this flips the sign to a positive total.
async function getUsageByUser(): Promise<Map<string, number>> {
  const map = new Map<string, number>();
  const db = await getDb();
  if (!db) return map;
  try {
    const rows = await db.collection("transactions").aggregate([
      { $group: { _id: "$user", total: { $sum: "$rawAmount" } } },
    ]).toArray();
    for (const r of rows) {
      if (r._id) map.set(String(r._id), Math.max(0, Math.round(-(r.total as number))));
    }
  } catch (e) {
    log(`getUsageByUser error: ${e}`);
  }
  return map;
}

async function deleteLcUser(email: string): Promise<void> {
  const db = await getDb();
  if (!db) throw new Error("MongoDB unavailable");
  await db.collection("users").deleteOne({ email });
}

async function setLcUserOrgRole(email: string, orgRole: "ADMIN" | "USER"): Promise<void> {
  const db = await getDb();
  if (!db) throw new Error("MongoDB unavailable");
  // Keep LibreChat's own `role` field in sync so Admin Settings link remains visible for ADMIN/OWNER
  const lcRole = orgRole === "USER" ? "USER" : "ADMIN";
  await db.collection("users").updateOne({ email }, { $set: { orgRole, role: lcRole } });
}

// ── Azure Table Storage ────────────────────────────────────────────────────────

const ACCOUNT = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY = process.env.AZURE_STORAGE_KEY ?? "";
const TABLE = process.env.TOKEN_TABLE_NAME || "agent365tokens";

// ── Auth resolution ─────────────────────────────────────────────────────────
//
// org-admin issues and verifies its OWN signed session tokens (below), backed
// by the same bcrypt password hash LibreChat itself uses to log the user in.
// This is a self-contained trust boundary: nobody can get a valid org_token
// without knowing that account's actual password, and nobody can forge one
// without knowing ORG_ADMIN_JWT_SECRET (a real Container App secret, never
// sent to the browser). There is NO path left that trusts a bare,
// caller-supplied email or an unsigned token — see feedback_2026-07-10
// security review for why those were removed.

const ORG_ADMIN_JWT_SECRET = process.env.JWT_SECRET ?? "";
const ORG_SESSION_ISSUER = "abk-org-admin";
const ORG_SESSION_TTL = "8h";

interface OrgSessionPayload {
  email: string;
  tier: Tier;
}

function signOrgSessionToken(email: string, tier: Tier): string {
  if (!ORG_ADMIN_JWT_SECRET) throw new Error("JWT_SECRET not configured");
  return jwt.sign({ email, tier }, ORG_ADMIN_JWT_SECRET, {
    issuer: ORG_SESSION_ISSUER,
    expiresIn: ORG_SESSION_TTL,
  });
}

function verifyOrgSessionToken(token: string): OrgSessionPayload | null {
  if (!ORG_ADMIN_JWT_SECRET || !token) return null;
  try {
    const payload = jwt.verify(token, ORG_ADMIN_JWT_SECRET, { issuer: ORG_SESSION_ISSUER }) as jwt.JwtPayload;
    if (!payload || typeof payload.email !== "string") return null;
    return { email: payload.email, tier: (payload.tier as Tier) ?? "USER" };
  } catch {
    return null;
  }
}

// ── org-admin's own SSO (PKCE authorization-code flow) ──────────────────────
//
// The OAuth `state` round-trip data used to live in a short-lived cookie
// (`abk_sso_<state>`). That works fine for a top-level sign-in, but org-admin
// is loaded in an iframe whose top-level document is LibreChat's origin —
// browsers classify ANY cookie set by agent365-bridge as cross-site for the
// whole redirect chain (start -> login.microsoftonline.com -> callback),
// regardless of SameSite value, because "site for cookies" is computed
// against the top-level document, not the iframe's own navigation. So the
// cookie never survived the round trip, `stateData` came back null on every
// attempt (even the very first silent one), and the callback fell straight
// into the "Sign-in failed" error page instead of silently succeeding or
// falling back to an interactive prompt.
//
// Fix: carry {verifier, silent, embed} inside the `state` value itself, as a
// short-lived signed JWT. AAD round-trips `state` verbatim via the URL query
// string, which isn't subject to cookie/storage partitioning at all — so
// this works identically inside a third-party iframe, with third-party
// cookies blocked, everywhere.

type Theme = "light" | "dark";

interface SsoStateData {
  verifier: string;
  silent: boolean;
  embed: boolean;
  theme: Theme;
}

function qtheme(req: Request): Theme {
  return qstr(req.query.theme) === "dark" ? "dark" : "light";
}

const SSO_STATE_ISSUER = "abk-org-sso-state";

function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function signSsoState(data: SsoStateData): string {
  return jwt.sign(data, ORG_ADMIN_JWT_SECRET, { issuer: SSO_STATE_ISSUER, expiresIn: "5m" });
}

function verifySsoState(state: string): SsoStateData | null {
  if (!ORG_ADMIN_JWT_SECRET || !state) return null;
  try {
    const payload = jwt.verify(state, ORG_ADMIN_JWT_SECRET, { issuer: SSO_STATE_ISSUER }) as jwt.JwtPayload;
    if (!payload || typeof payload.verifier !== "string") return null;
    return {
      verifier: payload.verifier,
      silent: !!payload.silent,
      embed: !!payload.embed,
      theme: payload.theme === "dark" ? "dark" : "light",
    };
  } catch {
    return null;
  }
}

async function resolveByM365Token(sessionToken: string): Promise<AuthInfo | null> {
  try {
    const cred = new AzureNamedKeyCredential(ACCOUNT, ACCOUNT_KEY);
    const client = new TableClient(`https://${ACCOUNT}.table.core.windows.net`, TABLE, cred);
    const rowKey = Buffer.from(sessionToken).toString("hex").substring(0, 256);
    const entity = await client.getEntity("tokens", rowKey);
    const email = entity["email"] as string | undefined;
    if (!email) return null;
    const user = await getLcUserByEmail(email);
    return { email, tier: getTier(email, user?.role ?? "USER", user?.orgRole) };
  } catch {
    return null;
  }
}

function qstr(v: unknown): string | undefined {
  if (typeof v === "string") return v;
  if (Array.isArray(v) && typeof v[0] === "string") return v[0];
  return undefined;
}

function bearerToken(req: Request): string | undefined {
  const h = req.headers.authorization;
  if (h && h.startsWith("Bearer ")) return h.slice(7).trim();
  return undefined;
}

async function resolveAuth(req: Request): Promise<AuthInfo | null> {
  const orgToken = bearerToken(req) ?? qstr(req.query.org_token);
  const m365Token = qstr(req.query.token);
  if (orgToken) {
    const session = verifyOrgSessionToken(orgToken);
    if (!session) return null;
    // Re-fetch the current role from Mongo rather than trusting the token's
    // snapshot, so a role change takes effect immediately without forcing
    // the affected user to log back in to org-admin.
    const user = await getLcUserByEmail(session.email);
    if (!user) return null;
    return { email: session.email, tier: getTier(session.email, user.role, user.orgRole) };
  }
  if (m365Token) return resolveByM365Token(m365Token);
  return null;
}

// ── HTML ─────────────────────────────────────────────────────────────────────

function buildHtml(email: string, tier: Tier, embed: boolean, theme: Theme): string {
  const authParamJs = `
const _p = new URLSearchParams(location.search);
const ORG_TOKEN  = _p.get('org_token') || '';
const M365_TOKEN = _p.get('token')     || '';
function authParam() {
  if (ORG_TOKEN)  return 'org_token=' + encodeURIComponent(ORG_TOKEN);
  return 'token=' + encodeURIComponent(M365_TOKEN);
}
// A token that just arrived via the SSO redirect — hand it to the parent
// LibreChat page so reopening Organization Settings in this tab doesn't
// need to repeat the SSO round-trip.
if (ORG_TOKEN) {
  try { window.parent.postMessage({ type: 'abk_org_session', token: ORG_TOKEN, email: '${email.replace(/'/g, "\\'")}' }, '*'); } catch(ex) {}
}
// Live theme sync — if the user toggles light/dark in LibreChat while this
// panel is already open, follow along without needing a reload.
window.addEventListener('message', function(evt) {
  if (evt && evt.data && evt.data.type === 'abk_theme') {
    document.documentElement.setAttribute('data-theme', evt.data.dark ? 'dark' : 'light');
  }
});`;

  const sidebarHtml = embed ? "" : `
<div class="sidebar">
  <div class="sb-section">
    <div class="sb-label">Settings</div>
    <a class="sb-item" href="#">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M8 8a3 3 0 100-6 3 3 0 000 6zm-5 6s-1 0-1-1 1-4 6-4 6 3 6 4-1 1-1 1H3z"/></svg>
      General
    </a>
  </div>
  <div class="sb-section">
    <div class="sb-label">People</div>
    <a class="sb-item active" href="#">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M15 14s1 0 1-1-1-4-5-4-5 3-5 4 1 1 1 1h8zm-9.978-1A.261.261 0 015 13c0-.366.268-1.14 1.004-1.844C6.717 10.48 7.742 10 9 10c.37 0 .724.043 1.05.12C9.5 10.48 9 11.25 9 12c0 .273.04.54.097.8H5.022zM4.5 8a2.5 2.5 0 100-5 2.5 2.5 0 000 5z"/></svg>
      Members
    </a>
  </div>
</div>`;

  const roleLabel = (t: Tier) => (t === "OWNER" ? "Primary owner" : t === "ADMIN" ? "Admin" : "User");

  return `<!DOCTYPE html>
<html lang="en" data-theme="${theme}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Organization Settings</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    @keyframes abkFadeUp{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes abkPulse{0%,100%{opacity:1}50%{opacity:.45}}
    /* ── Theme tokens — mirrors LibreChat's own light/dark palette (see patch.js) ── */
    :root{
      --bg-page:#fafafa;--surface:#fff;--border:#ececec;--border-soft:#f3f4f6;--border-input:#e5e7eb;--border-hover:#c9ccd1;--border-heavy:#d1d5db;
      --text-primary:#111;--text-secondary:#374151;--text-secondary-alt:#6b7280;--text-tertiary:#9ca3af;
      --accent:#0071bc;--accent-hover:#005a96;--accent-soft-bg:#eaf2fe;
      --danger:#dc2626;--danger-border:#fca5a5;--danger-soft-bg:#fef2f2;
      --success:#16a34a;--success-soft-bg:#f0fdf4;
      --hover-bg:#f9fafb;--row-hover-bg:#fafcff;
      --skeleton-a:#f0f0f0;--skeleton-b:#f6f6f6;
      --shadow-card:0 1px 2px rgba(16,24,40,.04);
      --shadow-modal:0 24px 60px rgba(0,0,0,.22);
      --overlay-bg:rgba(17,17,17,0.5);
    }
    :root[data-theme="dark"]{
      --bg-page:#101317;--surface:#15181d;--border:#282e35;--border-soft:#23282f;--border-input:#353c44;--border-hover:#454d57;--border-heavy:#454d57;
      --text-primary:#eef1f5;--text-secondary:#9aa5b1;--text-secondary-alt:#9aa5b1;--text-tertiary:#69737f;
      --accent:#2e97e0;--accent-hover:#4aa8e8;--accent-soft-bg:#10293b;
      --danger:#e46a6a;--danger-border:#8a3a3a;--danger-soft-bg:#2a1416;
      --success:#4ade80;--success-soft-bg:#0f2a17;
      --hover-bg:#1a1e24;--row-hover-bg:#181c22;
      --skeleton-a:#1e2229;--skeleton-b:#262b33;
      --shadow-card:0 1px 2px rgba(0,0,0,.3);
      --shadow-modal:0 24px 60px rgba(0,0,0,.55);
      --overlay-bg:rgba(0,0,0,.6);
    }
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg-page);color:var(--text-primary);min-height:100vh;font-size:14px;-webkit-font-smoothing:antialiased}
    a{text-decoration:none;color:inherit}
    .layout{display:flex;min-height:100vh;animation:abkFadeUp .25s ease}
    /* Sidebar */
    .sidebar{width:192px;flex-shrink:0;padding:16px 8px;border-right:1px solid var(--border);background:var(--surface)}
    .sb-section{margin-bottom:16px}
    .sb-label{font-size:11px;font-weight:500;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.05em;padding:0 8px;margin-bottom:2px}
    .sb-item{display:flex;align-items:center;padding:7px 8px;border-radius:7px;font-size:13px;color:var(--text-secondary);cursor:pointer;gap:8px;margin-bottom:1px;transition:background .15s ease,color .15s ease}
    .sb-item.active{background:var(--accent-soft-bg);color:var(--accent);font-weight:500}
    .sb-item:hover:not(.active){background:var(--hover-bg)}
    /* Main */
    .main{flex:1;padding:${embed ? "20px 24px" : "32px 40px"};max-width:900px}
    h1{font-size:18px;font-weight:600;color:var(--text-primary);margin-bottom:20px;letter-spacing:-.01em}
    /* Overview */
    .overview{display:flex;background:var(--surface);border:1px solid var(--border);border-radius:12px;margin-bottom:24px;overflow:hidden;box-shadow:var(--shadow-card)}
    .ov-item{flex:1;padding:16px 20px;border-right:1px solid var(--border-soft);transition:background .15s ease}
    .ov-item:last-child{border-right:none}
    .ov-label{font-size:11px;color:var(--text-tertiary);margin-bottom:5px;font-weight:500;text-transform:uppercase;letter-spacing:.05em}
    .ov-val{font-size:14px;font-weight:600;color:var(--text-primary)}
    .ov-sub{font-size:12px;color:var(--accent);margin-top:3px;font-weight:500}
    /* Members header */
    .m-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
    .m-title{font-size:15px;font-weight:600;color:var(--text-primary)}
    .btn-primary{background:var(--accent);color:#fff;border:none;border-radius:8px;padding:7px 15px;font-size:13px;font-weight:500;cursor:pointer;transition:background .15s ease,transform .1s ease,box-shadow .15s ease;box-shadow:0 1px 2px rgba(0,102,204,.25)}
    .btn-primary:hover{background:var(--accent-hover);box-shadow:0 2px 8px rgba(0,102,204,.3)}
    .btn-primary:active{transform:translateY(1px)}
    /* Table */
    .table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;box-shadow:var(--shadow-card)}
    .table-toolbar{display:flex;align-items:center;gap:10px;padding:10px 16px;border-bottom:1px solid var(--border-soft)}
    .search-wrap{position:relative;display:flex;align-items:center}
    .search-wrap svg{position:absolute;left:9px;pointer-events:none;color:var(--text-tertiary)}
    .search{border:1px solid var(--border-input);border-radius:7px;padding:6px 11px 6px 28px;font-size:13px;outline:none;width:200px;color:var(--text-primary);background:var(--surface);transition:border-color .15s ease,box-shadow .15s ease}
    .search:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(0,102,204,.12)}
    .count-lbl{font-size:13px;color:var(--text-tertiary)}
    table{width:100%;border-collapse:collapse}
    th{padding:10px 16px;text-align:left;font-size:11px;font-weight:500;color:var(--text-tertiary);border-bottom:1px solid var(--border);text-transform:uppercase;letter-spacing:.05em}
    td{padding:12px 16px;font-size:13px;border-bottom:1px solid var(--border-soft);vertical-align:middle}
    tr:last-child td{border-bottom:none}
    tr{transition:background .12s ease}
    tr:hover td{background:var(--row-hover-bg)}
    /* Name cell */
    .name-cell{display:flex;align-items:center;gap:10px}
    .avatar{width:30px;height:30px;border-radius:50%;color:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;flex-shrink:0;text-transform:uppercase;box-shadow:inset 0 0 0 1px rgba(255,255,255,.15)}
    .name-primary{font-size:13px;font-weight:500;color:var(--text-primary)}
    .name-secondary{font-size:12px;color:var(--text-tertiary)}
    .you-badge{font-size:11px;color:var(--accent);background:var(--accent-soft-bg);border-radius:5px;padding:1px 6px;margin-left:5px;font-weight:500}
    /* Role */
    .role-text{font-size:13px;color:var(--text-secondary)}
    .role-select{padding:4px 28px 4px 8px;border:1px solid var(--border-input);border-radius:7px;font-size:13px;cursor:pointer;background:var(--surface);color:var(--text-secondary);outline:none;appearance:auto;transition:border-color .15s ease}
    .role-select:hover{border-color:var(--border-hover)}
    .role-select:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(0,102,204,.12)}
    /* Actions */
    .btn-remove{background:none;border:none;cursor:pointer;padding:4px 10px;border-radius:6px;font-size:12.5px;color:var(--text-secondary-alt);white-space:nowrap;transition:background .15s ease,color .15s ease}
    .btn-remove:hover{background:var(--danger-soft-bg);color:var(--danger)}
    /* Error */
    .err-bar{background:var(--danger-soft-bg);border:1px solid var(--danger-border);border-radius:9px;padding:10px 14px;color:var(--danger);font-size:13px;margin-bottom:16px;display:none;animation:abkFadeUp .2s ease}
    .empty-cell{text-align:center;padding:56px 20px;color:var(--text-tertiary);font-size:13px}
    .empty-cell svg{display:block;margin:0 auto 12px;color:var(--border-heavy)}
    .skeleton{background:linear-gradient(90deg,var(--skeleton-a),var(--skeleton-b),var(--skeleton-a));border-radius:6px;animation:abkPulse 1.3s ease infinite}
    /* Add member modal */
    .add-input{width:100%;padding:9px 12px;border:1px solid var(--border-input);border-radius:8px;font-size:13px;margin-bottom:10px;outline:none;box-sizing:border-box;transition:border-color .15s ease,box-shadow .15s ease;background:var(--surface);color:var(--text-primary)}
    .add-input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(0,102,204,.12)}
    #addMsg{font-size:13px;padding:10px 12px;border-radius:8px;margin-bottom:12px;display:none}
  </style>
</head>
<body>
<div class="layout">
${sidebarHtml}
  <div class="main">
    <h1>Organization settings</h1>
    <div id="errBox" class="err-bar"></div>
    <div class="overview">
      <div class="ov-item">
        <div class="ov-label">Allowed email domains</div>
        <div class="ov-val">abk.gr</div>
      </div>
      <div class="ov-item">
        <div class="ov-label">Total members</div>
        <div class="ov-val" id="totalMembers">—</div>
      </div>
      <div class="ov-item">
        <div class="ov-label">Signed in as</div>
        <div class="ov-val">${email}</div>
        <div class="ov-sub">${roleLabel(tier)}</div>
      </div>
    </div>
    <div class="m-header">
      <div class="m-title">Members</div>
      ${tier !== "USER" ? '<button class="btn-primary" onclick="showAddModal()">+ Add member</button>' : ""}
    </div>
    <div class="table-wrap">
      <div class="table-toolbar">
        <div class="search-wrap">
          <svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="7" cy="7" r="5.2"/><path d="M11 11l3.5 3.5" stroke-linecap="round"/></svg>
          <input class="search" id="search" type="text" placeholder="Search…" oninput="filterUsers()">
        </div>
        <span class="count-lbl" id="countLbl"></span>
      </div>
      <table>
        <thead><tr>
          <th>Name</th>
          <th>Role</th>
          <th>Tokens used</th>
          ${tier !== "USER" ? "<th></th>" : ""}
        </tr></thead>
        <tbody id="tbody">${[0, 1, 2].map(() => `<tr><td><div class="name-cell"><div class="skeleton" style="width:30px;height:30px;border-radius:50%"></div><div style="flex:1"><div class="skeleton" style="width:120px;height:11px;margin-bottom:6px"></div><div class="skeleton" style="width:160px;height:10px"></div></div></div></td><td><div class="skeleton" style="width:60px;height:11px"></div></td><td><div class="skeleton" style="width:50px;height:11px"></div></td>${tier !== "USER" ? "<td></td>" : ""}</tr>`).join("")}</tbody>
      </table>
    </div>
  </div>
</div>
<!-- Add member modal -->
<div id="addModal" style="display:none;position:fixed;inset:0;z-index:9999;background:var(--overlay-bg);align-items:center;justify-content:center;backdrop-filter:blur(2px);">
  <div style="background:var(--surface);border-radius:14px;padding:28px 32px;width:380px;max-width:95vw;box-shadow:var(--shadow-modal);animation:abkFadeUp .18s ease;">
    <h2 style="font-size:15px;font-weight:600;color:var(--text-primary);margin-bottom:6px">Add member</h2>
    <p style="font-size:13px;color:var(--text-secondary-alt);margin-bottom:18px">Invite a new @abk.gr member. They'll sign in with their ABK Microsoft 365 account — no password to set.</p>
    <div id="addMsg"></div>
    <input id="addEmail" class="add-input" type="email" placeholder="you@abk.gr" style="margin-bottom:16px">
    <button id="addSubmitBtn" class="btn-primary" onclick="submitAdd()" style="width:100%;padding:10px;margin-bottom:8px">Send invitation</button>
    <button onclick="hideAddModal()" style="width:100%;padding:8px;background:none;border:none;font-size:13px;color:var(--text-secondary-alt);cursor:pointer;border-radius:8px;transition:background .15s ease" onmouseover="this.style.background='var(--hover-bg)'" onmouseout="this.style.background='none'">Cancel</button>
  </div>
</div>
<script>
const MY_TIER = '${tier}';
const MY_EMAIL = '${email}';
${authParamJs}
let allUsers = [];

const AV_COLORS = [['#8b5cf6','#6d28d9'],['#3b82f6','#1d4ed8'],['#10b981','#047857'],['#f59e0b','#b45309'],['#ef4444','#b91c1c'],['#ec4899','#be185d'],['#06b6d4','#0e7490'],['#84cc16','#4d7c0f'],['#a855f7','#7e22ce'],['#0ea5e9','#0369a1']];
function avColor(n) { var h=0; for(var i=0;i<n.length;i++) h=(h*31+n.charCodeAt(i))&0x7fffffff; var c=AV_COLORS[h%AV_COLORS.length]; return 'linear-gradient(135deg,'+c[0]+','+c[1]+')'; }
function initials(n) { var p=n.trim().split(/\\s+/); return (p.length>=2?p[0][0]+p[1][0]:n.slice(0,2)).toUpperCase(); }
function roleLabel(t) { return t==='OWNER'?'Primary owner':t==='ADMIN'?'Admin':'User'; }
function canActOn(t) { return (MY_TIER==='OWNER' || MY_TIER==='ADMIN') && t !== 'OWNER'; }

async function load() {
  try {
    const r = await fetch('/org-admin/api/users?' + authParam());
    if (!r.ok) throw new Error(await r.text());
    const d = await r.json();
    allUsers = d.users || [];
    render(allUsers);
    document.getElementById('totalMembers').textContent = allUsers.length;
    document.getElementById('countLbl').textContent = allUsers.length + ' member' + (allUsers.length!==1?'s':'');
    // Tell parent LibreChat page the current user's role (for Account settings display)
    try { window.parent.postMessage({ type: 'abk_org_role', role: MY_TIER }, '*'); } catch(ex) {}
  } catch(e) { showErr('Failed to load: ' + e.message); }
}

function render(users) {
  const tb = document.getElementById('tbody');
  if (!users.length) { tb.innerHTML = '<tr><td class="empty-cell" colspan="4"><svg width="32" height="32" viewBox="0 0 16 16" fill="currentColor"><path d="M15 14s1 0 1-1-1-4-5-4-5 3-5 4 1 1 1 1h8zm-9.978-1A.261.261 0 015 13c0-.366.268-1.14 1.004-1.844C6.717 10.48 7.742 10 9 10c.37 0 .724.043 1.05.12C9.5 10.48 9 11.25 9 12c0 .273.04.54.097.8H5.022zM4.5 8a2.5 2.5 0 100-5 2.5 2.5 0 000 5z"/></svg>No members found</td></tr>'; return; }
  tb.innerHTML = users.map(u => {
    const name = u.name || u.email.split('@')[0];
    const isMe = u.email === MY_EMAIL;
    const canAct = !isMe && canActOn(u.tier);
    const color = avColor(name);
    const youBadge = isMe ? '<span class="you-badge">You</span>' : '';

    let roleCell;
    const canChangeRole = !isMe && u.tier !== 'OWNER' && MY_TIER !== 'USER';
    if (u.tier === 'OWNER') {
      roleCell = '<span class="role-text">Primary owner</span>';
    } else if (canChangeRole) {
      roleCell = \`<select class="role-select" onchange="changeRole('\${esc(u.email)}',this.value)">
        <option \${u.tier==='ADMIN'?'selected':''} value="ADMIN">Admin</option>
        <option \${u.tier==='USER'?'selected':''} value="USER">User</option>
      </select>\`;
    } else {
      roleCell = \`<span class="role-text">\${roleLabel(u.tier)}</span>\`;
    }

    const actCell = MY_TIER !== 'USER' ? \`<td style="text-align:right">
      \${canAct ? \`<button class="btn-remove" onclick="delUser('\${esc(u.email)}')">Remove</button>\` : ''}
    </td>\` : '';

    const usageCell = \`<span class="role-text">\${(u.consumed || 0).toLocaleString('en-US')}</span>\`;

    return \`<tr>
      <td><div class="name-cell">
        <div class="avatar" style="background:\${color}">\${initials(name)}</div>
        <div><div class="name-primary">\${name}\${youBadge}</div><div class="name-secondary">\${u.email}</div></div>
      </div></td>
      <td>\${roleCell}</td>
      <td>\${usageCell}</td>
      \${actCell}
    </tr>\`;
  }).join('');
}

function filterUsers() {
  const q = document.getElementById('search').value.toLowerCase();
  const f = q ? allUsers.filter(u=>u.email.toLowerCase().includes(q)||(u.name||'').toLowerCase().includes(q)) : allUsers;
  render(f);
  document.getElementById('countLbl').textContent = f.length + ' member' + (f.length!==1?'s':'');
}

async function delUser(email) {
  if (!confirm('Remove ' + email + ' from the organization?')) return;
  const r = await fetch('/org-admin/api/users/'+encodeURIComponent(email)+'?'+authParam(), {method:'DELETE'});
  if (r.ok) {
    allUsers = allUsers.filter(u=>u.email!==email);
    render(allUsers);
    document.getElementById('totalMembers').textContent = allUsers.length;
    document.getElementById('countLbl').textContent = allUsers.length + ' member' + (allUsers.length!==1?'s':'');
  } else showErr('Error: ' + await r.text());
}

async function changeRole(email, role) {
  const r = await fetch('/org-admin/api/users/'+encodeURIComponent(email)+'/role?'+authParam(), {
    method:'PATCH', headers:{'Content-Type':'application/json'}, body:JSON.stringify({role})
  });
  if (!r.ok) { showErr('Error: ' + await r.text()); load(); }
  else {
    const u = allUsers.find(x=>x.email===email);
    if (u) { u.tier = role; render(allUsers); }
  }
}

function showErr(msg) {
  const el=document.getElementById('errBox');
  el.textContent=msg; el.style.display='block';
  setTimeout(()=>el.style.display='none',6000);
}
function esc(s) { return s.replace(/'/g,"\\'"); }

function showAddModal() {
  document.getElementById('addEmail').value='';
  var msg=document.getElementById('addMsg'); msg.style.display='none';
  var btn=document.getElementById('addSubmitBtn'); btn.disabled=false; btn.textContent='Send invitation'; btn.style.display='';
  var m=document.getElementById('addModal'); m.style.display='flex';
}
function hideAddModal() { document.getElementById('addModal').style.display='none'; }
document.getElementById('addModal').addEventListener('click', function(e){ if(e.target===this) hideAddModal(); });

async function submitAdd() {
  var email=document.getElementById('addEmail').value.trim().toLowerCase();
  var msg=document.getElementById('addMsg');
  function showAddMsg(txt,isErr){
    msg.style.display='block';
    msg.style.background=isErr?'var(--danger-soft-bg)':'var(--success-soft-bg)';
    msg.style.color=isErr?'var(--danger)':'var(--success)';
    msg.textContent=txt;
  }
  if(!email||email.indexOf('@')===-1){showAddMsg('Enter a valid email.',true);return;}
  var btn=document.getElementById('addSubmitBtn');
  btn.disabled=true; btn.textContent='Sending…';
  try {
    const r=await fetch('/org-admin/api/users?'+authParam(),{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({email})
    });
    const d=await r.json();
    if(r.ok){
      showAddMsg('Invitation sent!',false);
      btn.style.display='none';
      setTimeout(function(){ hideAddModal(); load(); },1200);
    } else {
      showAddMsg(d.error||'Error sending invitation.',true);
      btn.disabled=false; btn.textContent='Send invitation';
    }
  } catch(e){ showAddMsg('Connection error.',true); btn.disabled=false; btn.textContent='Send invitation'; }
}

load();
</script>
</body>
</html>`;
}

// ── SSO error / status pages (shown in embed mode for the SSO flow) ─────────

function buildSsoErrorPage(message: string, opts?: { retry?: boolean; popupClose?: boolean; theme?: Theme }): string {
  const theme: Theme = opts?.theme ?? "light";
  const actionBtn = opts?.popupClose
    ? `<button onclick="window.close()">Close window</button>`
    : opts?.retry
    ? `<button onclick="location.href='/org-admin/api/sso/start?silent=0&embed=1&theme=${theme}'">Try again</button>`
    : "";
  return `<!DOCTYPE html>
<html lang="en" data-theme="${theme}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Organization Settings</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    @keyframes abkFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    :root{--bg-page:#fafafa;--surface:#fff;--border:#f0f0f0;--text-primary:#111;--text-secondary-alt:#6b7280;--accent:#0071bc;--accent-hover:#0055bb;--shadow-modal:0 8px 28px rgba(16,24,40,.08)}
    :root[data-theme="dark"]{--bg-page:#101317;--surface:#15181d;--border:#282e35;--text-primary:#eef1f5;--text-secondary-alt:#9aa5b1;--accent:#2e97e0;--accent-hover:#4aa8e8;--shadow-modal:0 8px 28px rgba(0,0,0,.4)}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg-page);min-height:100vh;display:flex;align-items:center;justify-content:center}
    .card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:36px 40px;box-shadow:var(--shadow-modal);width:320px;animation:abkFadeUp .25s ease;text-align:center}
    h2{font-size:16px;font-weight:600;color:var(--text-primary);margin-bottom:8px}
    p{color:var(--text-secondary-alt);font-size:13px;margin-bottom:20px;line-height:1.5}
    button{width:100%;padding:11px;background:var(--accent);color:#fff;border:none;border-radius:9px;font-size:14px;font-weight:500;cursor:pointer;transition:background .15s ease,box-shadow .15s ease}
    button:hover{background:var(--accent-hover);box-shadow:0 2px 10px rgba(0,102,204,.28)}
  </style>
</head>
<body>
  <div class="card">
    <h2>Sign-in required</h2>
    <p>${message.replace(/</g, "&lt;")}</p>
    ${actionBtn}
  </div>
</body>
</html>`;
}

// ── SSO popup helpers ─────────────────────────────────────────────────────────
//
// Microsoft's own sign-in page sends X-Frame-Options and refuses to render
// inside ANY iframe (clickjacking protection on their end, not something we
// can turn off) — silent (prompt=none) attempts never hit this because they
// never render visible UI, they just redirect. But the moment silent SSO
// fails and we fall back to a real interactive sign-in, redirecting the
// iframe itself straight to Microsoft gets silently blocked by the browser,
// which is exactly the "blocked content" icon instead of a login page.
// Fix: for interactive sign-in, open Microsoft's page in a real pop-up
// window instead (pop-ups are top-level browsing contexts, not framed, so
// Microsoft is happy to render there) and hand the resulting session back to
// the iframe via postMessage once it completes.

function buildPopupLauncherPage(authUrl: string, theme: Theme): string {
  return `<!DOCTYPE html>
<html lang="en" data-theme="${theme}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Organization Settings</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    @keyframes abkFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    :root{--bg-page:#fafafa;--surface:#fff;--border:#f0f0f0;--text-primary:#111;--text-secondary-alt:#6b7280;--accent:#0071bc;--accent-hover:#0055bb;--shadow-modal:0 8px 28px rgba(16,24,40,.08)}
    :root[data-theme="dark"]{--bg-page:#101317;--surface:#15181d;--border:#282e35;--text-primary:#eef1f5;--text-secondary-alt:#9aa5b1;--accent:#2e97e0;--accent-hover:#4aa8e8;--shadow-modal:0 8px 28px rgba(0,0,0,.4)}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg-page);min-height:100vh;display:flex;align-items:center;justify-content:center}
    .card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:36px 40px;box-shadow:var(--shadow-modal);width:320px;animation:abkFadeUp .25s ease;text-align:center}
    h2{font-size:16px;font-weight:600;color:var(--text-primary);margin-bottom:8px}
    p{color:var(--text-secondary-alt);font-size:13px;margin-bottom:20px;line-height:1.5}
    a.btn{display:block;width:100%;padding:11px;background:var(--accent);color:#fff;border:none;border-radius:9px;font-size:14px;font-weight:500;cursor:pointer;transition:background .15s ease;text-decoration:none;box-sizing:border-box}
    a.btn:hover{background:var(--accent-hover)}
  </style>
</head>
<body>
  <div class="card">
    <h2>Sign-in required</h2>
    <p>A Microsoft sign-in window should have opened. If you don't see it, your browser may have blocked the pop-up.</p>
    <a class="btn" id="openBtn" href="${authUrl}" target="abk_sso_popup">Open sign-in window</a>
  </div>
  <script>
    window.addEventListener('message', function(evt) {
      if (evt && evt.data && evt.data.type === 'abk_org_session' && evt.data.token) {
        location.href = '/org-admin?embed=1&org_token=' + encodeURIComponent(evt.data.token) + '&theme=${theme}';
      }
    });
    try { window.open(${JSON.stringify(authUrl)}, 'abk_sso_popup', 'width=480,height=640'); } catch(ex) {}
  </script>
</body>
</html>`;
}

function buildPopupSuccessPage(token: string, email: string, theme: Theme): string {
  return `<!DOCTYPE html>
<html lang="en" data-theme="${theme}">
<head>
  <meta charset="utf-8">
  <title>Signed in</title>
  <style>
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:${theme === "dark" ? "#101317" : "#fafafa"};color:${theme === "dark" ? "#eef1f5" : "#111"};display:flex;align-items:center;justify-content:center;min-height:100vh;font-size:14px}
  </style>
</head>
<body>
  <p>Signed in — you can close this window.</p>
  <script>
    try {
      if (window.opener) {
        window.opener.postMessage({ type: 'abk_org_session', token: ${JSON.stringify(token)}, email: ${JSON.stringify(email)} }, '*');
      }
    } catch(ex) {}
    window.close();
  </script>
</body>
</html>`;
}

// ── Password Reset Helpers ────────────────────────────────────────────────────

function buildResetPage(token: string): string {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Reset Password — ABK Assistant</title>
<style>
  @keyframes abkFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#fafafa;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
  .card{background:#fff;border:1px solid #f0f0f0;border-radius:16px;padding:36px;width:340px;box-shadow:0 12px 36px rgba(16,24,40,.1);animation:abkFadeUp .25s ease;}
  h2{font-size:18px;font-weight:600;margin:0 0 6px;color:#111;}
  p{font-size:13px;color:#6b7280;margin:0 0 20px;line-height:1.5}
  input{width:100%;padding:10px 14px;border:1px solid #e3e3e3;border-radius:9px;font-size:14px;margin-bottom:10px;box-sizing:border-box;outline:none;transition:border-color .15s ease,box-shadow .15s ease}
  input:focus{border-color:#0071BC;box-shadow:0 0 0 3px rgba(0,102,204,.14);}
  button{width:100%;padding:11px;background:#0071BC;color:#fff;border:none;border-radius:9px;font-size:14px;font-weight:500;cursor:pointer;transition:background .15s ease,box-shadow .15s ease}
  button:hover{background:#005a96;box-shadow:0 2px 10px rgba(0,102,204,.28);}
  #msg{display:none;padding:10px 12px;border-radius:8px;font-size:13px;margin-bottom:12px;}
  img{height:36px;margin-bottom:20px;display:block;}
</style></head>
<body><div class="card">
  <img src="${SERVER_BASE_URL}/assets/abk-logo.png" alt="ABK Solutions">
  <h2>Set New Password</h2>
  <p>Choose a new password for your ABK Assistant account.</p>
  <div id="msg"></div>
  <input id="pw1" type="password" placeholder="New password (min 8 chars)">
  <input id="pw2" type="password" placeholder="Confirm new password" style="margin-bottom:16px;">
  <button onclick="submit()">Set Password</button>
</div>
<script>
function submit(){
  var pw1=document.getElementById('pw1').value;
  var pw2=document.getElementById('pw2').value;
  var msg=document.getElementById('msg');
  if(pw1.length<8){show('Password must be at least 8 characters.',true);return;}
  if(pw1!==pw2){show('Passwords do not match.',true);return;}
  fetch('${SERVER_BASE_URL}/org-admin/api/reset-password-confirm',{
    method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({token:'${token}',newPassword:pw1})
  }).then(function(r){return r.json();}).then(function(d){
    if(d.success){
      document.querySelector('.card').innerHTML='<img src="${SERVER_BASE_URL}/assets/abk-logo.png" alt="ABK" style="height:36px;margin-bottom:20px;display:block;"><h2 style="color:#0071BC">Password Updated!</h2><p>Your password has been changed. You can now <a href="${SERVER_BASE_URL.replace("agent365-bridge", "abkagent-backup").replace(/agent365.*/, "")}" style="color:#0071BC">log in</a>.</p>';
    } else { show(d.error||'Something went wrong.',true); }
  }).catch(function(){show('Connection error.',true);});
}
function show(txt,err){var m=document.getElementById('msg');m.style.display='block';m.style.background=err?'#fef2f2':'#f0fdf4';m.style.color=err?'#dc2626':'#16a34a';m.textContent=txt;}
document.addEventListener('keydown',function(e){if(e.key==='Enter')submit();});
</script></body></html>`;
}

function buildResetErrorPage(message: string): string {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Reset Password — ABK Assistant</title>
<style>
  @keyframes abkFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#fafafa;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
  .card{background:#fff;border:1px solid #f0f0f0;border-radius:16px;padding:36px;width:340px;box-shadow:0 12px 36px rgba(16,24,40,.1);animation:abkFadeUp .25s ease;}
  img{height:36px;margin-bottom:20px;display:block;}
</style>
</head><body><div class="card">
  <img src="${SERVER_BASE_URL}/assets/abk-logo.png" alt="ABK Solutions">
  <h2 style="color:#dc2626;font-size:18px;">Link Invalid</h2>
  <p style="color:#6b7280;font-size:13px;line-height:1.5">${message}</p>
</div></body></html>`;
}

async function sendResetEmail(toEmail: string, resetUrl: string, graphToken: string): Promise<boolean> {
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(PRIMARY_OWNER_EMAIL)}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: "ABK Assistant — Password Reset",
          body: {
            contentType: "HTML",
            content: `<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;">
                <img src="${SERVER_BASE_URL}/assets/abk-logo.png" height="36" style="margin-bottom:24px;display:block;">
                <h2 style="color:#0071BC;font-size:20px;margin:0 0 12px;">Reset your password</h2>
                <p style="color:#444;line-height:1.6;">We received a request to reset your <strong>ABK Assistant</strong> password. Click below to set a new password. This link expires in <strong>1 hour</strong>.</p>
                <a href="${resetUrl}" style="display:inline-block;margin:24px 0;padding:12px 28px;background:#0071BC;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;">Reset Password</a>
                <p style="color:#888;font-size:12px;">If you didn't request this, you can ignore this email.</p>
              </div>`,
          },
          toRecipients: [{ emailAddress: { address: toEmail } }],
        },
        saveToSentItems: false,
      }),
    });
    if (!res.ok) {
      log(`Graph sendMail failed: ${await res.text()}`);
      return false;
    }
    return true;
  } catch (e) {
    log(`sendResetEmail error: ${e}`);
    return false;
  }
}

async function sendInviteEmail(toEmail: string, graphToken: string): Promise<boolean> {
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(PRIMARY_OWNER_EMAIL)}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject: "You've been invited to ABK Assistant",
          body: {
            contentType: "HTML",
            content: `<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;">
                <img src="${SERVER_BASE_URL}/assets/abk-logo.png" height="36" style="margin-bottom:24px;display:block;">
                <h2 style="color:#0071BC;font-size:20px;margin:0 0 12px;">You're invited to ABK Assistant</h2>
                <p style="color:#444;line-height:1.6;">An administrator has added you to ABK Assistant. Sign in with your ABK Microsoft 365 account to get started — no separate password needed.</p>
                <a href="${LIBRECHAT_URL}/login" style="display:inline-block;margin:24px 0;padding:12px 28px;background:#0071BC;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;">Sign in to ABK Assistant</a>
                <p style="color:#888;font-size:12px;">If you weren't expecting this, you can ignore this email.</p>
              </div>`,
          },
          toRecipients: [{ emailAddress: { address: toEmail } }],
        },
        saveToSentItems: false,
      }),
    });
    if (!res.ok) {
      log(`Graph sendMail (invite) failed: ${await res.text()}`);
      return false;
    }
    return true;
  } catch (e) {
    log(`sendInviteEmail error: ${e}`);
    return false;
  }
}

// ── CORS (scoped to /org-admin only) ─────────────────────────────────────────
//
// mcp-proxy-server.ts applies a wildcard Access-Control-Allow-Origin to the
// whole app (reasonable for the bearer-token MCP/OAuth routes it's meant
// for). org-admin's routes carry real session tokens and password data, so
// they get a real origin allowlist here instead, which — since it runs
// after and re-sets the header — overrides that permissive default for
// every /org-admin* route specifically.
const ORG_ADMIN_ALLOWED_ORIGIN_RE = process.env.ORG_ADMIN_ALLOWED_ORIGINS
  ? new RegExp(
      "^(" +
        process.env.ORG_ADMIN_ALLOWED_ORIGINS.split(",")
          .map((s) => s.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("|") +
        ")$"
    )
  // Azure Container Apps FQDNs have multiple dot-separated labels before the
  // suffix (e.g. abkagent-backup.mangoriver-ca6f548b.swedencentral.
  // azurecontainerapps.io) — the previous version of this regex only
  // allowed a single label and silently rejected every real deployment,
  // stripping the CORS header on legitimate cross-origin calls (patch.js's
  // "Forgot password?" fetch) so the browser blocked reading the response
  // even though the request succeeded server-side (misleading "Connection
  // error" while the reset email/password change actually went through).
  : /^https:\/\/[a-z0-9-]+(\.[a-z0-9-]+)*\.azurecontainerapps\.io$/i;

function orgAdminCors(req: Request, res: Response, next: () => void): void {
  const origin = req.headers.origin;
  if (origin && ORG_ADMIN_ALLOWED_ORIGIN_RE.test(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  } else {
    res.removeHeader("Access-Control-Allow-Origin");
  }
  if (req.method === "OPTIONS") {
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, PATCH, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    res.sendStatus(204);
    return;
  }
  next();
}

// ── Express endpoints ──────────────────────────────────────────────────────────

export function registerOrgAdminEndpoints(app: Express): void {
  app.use("/org-admin", orgAdminCors);

  app.get("/org-admin", async (req: Request, res: Response) => {
    res.removeHeader("X-Frame-Options");
    const embed = qstr(req.query.embed) === "1";
    const theme = qtheme(req);
    const info = await resolveAuth(req);
    if (!info) {
      if (embed) {
        // No valid session yet — kick off org-admin's own SSO flow (silent
        // first: if this browser already has an active Microsoft session
        // from the LibreChat login just now, this resolves invisibly with
        // no prompt at all). See /org-admin/api/sso/start.
        res.redirect(`/org-admin/api/sso/start?silent=1&embed=1&theme=${theme}`);
      } else {
        res.status(401).send("Unauthorized. Open via ABK Assistant.");
      }
      return;
    }
    res.setHeader("Content-Type", "text/html");
    res.send(buildHtml(info.email, info.tier, embed, theme));
  });

  // ── org-admin's own Microsoft SSO (separate origin from LibreChat, so it
  // needs its own real, verifiable sign-in rather than trusting anything the
  // parent page claims — see the comment above ORG_SSO_CLIENT_ID). ──────────

  app.get("/org-admin/api/sso/start", (req: Request, res: Response) => {
    const theme = qtheme(req);
    if (!ORG_SSO_CLIENT_ID || !ORG_SSO_CLIENT_SECRET || !ORG_SSO_TENANT_ID) {
      res.status(500).send(buildSsoErrorPage("SSO is not configured for Organization Settings yet.", { theme }));
      return;
    }
    const silent = qstr(req.query.silent) !== "0";
    const embed = qstr(req.query.embed) === "1";
    const verifier = base64url(crypto.randomBytes(32));
    const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
    const state = signSsoState({ verifier, silent, embed, theme });
    const params = new URLSearchParams({
      client_id: ORG_SSO_CLIENT_ID,
      response_type: "code",
      redirect_uri: ORG_SSO_REDIRECT_URI,
      response_mode: "query",
      scope: "openid profile email",
      state,
      code_challenge: challenge,
      code_challenge_method: "S256",
    });
    if (silent) params.set("prompt", "none");
    const authUrl = `https://login.microsoftonline.com/${ORG_SSO_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
    if (silent) {
      // No visible UI is ever rendered for prompt=none, so redirecting the
      // iframe itself straight there is safe.
      res.redirect(authUrl);
    } else {
      // Interactive sign-in DOES render Microsoft's own UI, which refuses to
      // be framed — hand it off to a pop-up instead (see buildPopupLauncherPage).
      res.send(buildPopupLauncherPage(authUrl, theme));
    }
  });

  app.get("/org-admin/api/sso/callback", async (req: Request, res: Response) => {
    const state = qstr(req.query.state) ?? "";
    const stateData = verifySsoState(state);

    const error = qstr(req.query.error);
    const fallbackTheme = qtheme(req);
    if (error || !stateData) {
      // Silent attempts routinely "fail" this way when the browser has no
      // active Microsoft session yet (interaction_required / login_required)
      // — that's expected, not a real error, so fall through to a real
      // interactive sign-in instead of showing the user anything.
      if (stateData?.silent) {
        res.redirect(`/org-admin/api/sso/start?silent=0&embed=${stateData.embed ? "1" : "0"}&theme=${stateData.theme}`);
        return;
      }
      // Reaching here means this wasn't (or we can't tell it was) a silent
      // attempt — silent failures are always handled above, so by this point
      // we're inside the interactive pop-up, not the iframe.
      res.status(401).send(
        buildSsoErrorPage(qstr(req.query.error_description) || error || "Sign-in failed.", {
          popupClose: true,
          theme: stateData?.theme ?? fallbackTheme,
        })
      );
      return;
    }

    // From here on, stateData.silent reliably tells us whether we're still
    // running inside the iframe (silent) or inside the pop-up (interactive)
    // — see buildPopupLauncherPage/buildPopupSuccessPage.
    const errOpts = (): { retry?: boolean; popupClose?: boolean; theme: Theme } =>
      stateData.silent ? { retry: true, theme: stateData.theme } : { popupClose: true, theme: stateData.theme };

    const code = qstr(req.query.code);
    if (!code) {
      res.status(400).send(buildSsoErrorPage("Missing authorization code.", errOpts()));
      return;
    }

    try {
      const tokenRes = await fetch(`https://login.microsoftonline.com/${ORG_SSO_TENANT_ID}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: ORG_SSO_CLIENT_ID,
          client_secret: ORG_SSO_CLIENT_SECRET,
          grant_type: "authorization_code",
          code,
          redirect_uri: ORG_SSO_REDIRECT_URI,
          code_verifier: stateData.verifier,
        }),
      });
      if (!tokenRes.ok) {
        log(`sso token exchange failed: ${await tokenRes.text()}`);
        if (stateData.silent) {
          res.redirect(`/org-admin/api/sso/start?silent=0&embed=${stateData.embed ? "1" : "0"}&theme=${stateData.theme}`);
          return;
        }
        res.status(401).send(buildSsoErrorPage("Could not complete sign-in.", errOpts()));
        return;
      }
      const tokenData = (await tokenRes.json()) as { access_token?: string };
      if (!tokenData.access_token) {
        res.status(401).send(buildSsoErrorPage("Could not complete sign-in.", errOpts()));
        return;
      }
      const userinfoRes = await fetch("https://graph.microsoft.com/oidc/userinfo", {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });
      if (!userinfoRes.ok) {
        res.status(401).send(buildSsoErrorPage("Could not verify your identity.", errOpts()));
        return;
      }
      const profile = (await userinfoRes.json()) as { email?: string; preferred_username?: string };
      const email = (profile.email || profile.preferred_username || "").toLowerCase().trim();
      if (!email) {
        res.status(401).send(buildSsoErrorPage("Microsoft did not return an email address.", errOpts()));
        return;
      }
      const user = await getLcUserByEmail(email);
      if (!user) {
        res.status(403).send(
          buildSsoErrorPage(`${email} isn't set up in ABK Assistant yet. Ask an administrator to add you first.`, errOpts())
        );
        return;
      }
      const tier = getTier(email, user.role, user.orgRole);
      const token = signOrgSessionToken(email, tier);
      log(`org-admin SSO login: ${email} (${tier}) [${stateData.silent ? "silent" : "interactive"}]`);
      if (!stateData.silent) {
        // Interactive flow ran in the pop-up — hand the token back to the
        // iframe (window.opener) and let the pop-up close itself.
        res.send(buildPopupSuccessPage(token, email, stateData.theme));
        return;
      }
      res.redirect(
        `/org-admin?embed=${stateData.embed ? "1" : "0"}&org_token=${encodeURIComponent(token)}&theme=${stateData.theme}`
      );
    } catch (e) {
      log(`/api/sso/callback error: ${e}`);
      if (stateData?.silent) {
        res.redirect(`/org-admin/api/sso/start?silent=0&embed=${stateData.embed ? "1" : "0"}&theme=${stateData.theme}`);
        return;
      }
      res.status(500).send(buildSsoErrorPage("Server error during sign-in.", { popupClose: true, theme: stateData?.theme ?? fallbackTheme }));
    }
  });

  app.get("/org-admin/api/users", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    try {
      const [lcUsers, usage] = await Promise.all([getLcUsers(), getUsageByUser()]);
      const users = lcUsers.map((u) => ({
        email: u.email,
        name: u.name || u.username || u.email.split("@")[0],
        tier: getTier(u.email, u.role, u.orgRole),
        consumed: usage.get(String(u._id)) ?? 0,
      }));
      res.json({ users, tier: info.tier });
    } catch (e) {
      log(`/api/users error: ${e}`);
      res.status(500).json({ error: String(e) });
    }
  });

  app.post("/org-admin/api/users", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info || info.tier === "USER") {
      res.status(403).json({ error: "Forbidden" });
      return;
    }
    const { email } = req.body ?? {};
    if (!email) {
      res.status(400).json({ error: "email is required" });
      return;
    }
    const normalizedEmail = String(email).toLowerCase().trim();
    if (!normalizedEmail.endsWith("@abk.gr")) {
      res.status(400).json({ error: "Only @abk.gr email addresses are allowed" });
      return;
    }
    const db = await getDb();
    if (!db) {
      res.status(500).json({ error: "Database unavailable" });
      return;
    }
    const existing = await getLcUserByEmail(normalizedEmail);
    if (existing) {
      res.status(409).json({ error: "An account with this email already exists" });
      return;
    }
    try {
      const username = normalizedEmail.split("@")[0].replace(/[^a-zA-Z0-9_-]/g, "").toLowerCase();
      // No password: this account is pre-approved for SSO only. LibreChat's
      // openidStrategy matches by email + provider === "openid" and fills in
      // openidId itself on the member's first real Microsoft sign-in.
      await db.collection("users").insertOne({
        name: username,
        username,
        email: normalizedEmail,
        provider: "openid",
        role: "USER",
        orgRole: "USER",
        emailVerified: true,
        createdAt: new Date(),
        updatedAt: new Date(),
        __v: 0,
      });
      log(`${info.tier} ${info.email} invited ${normalizedEmail}`);

      void (async () => {
        const adminSession = await getSessionTokenByEmail(PRIMARY_OWNER_EMAIL);
        if (!adminSession) {
          log(`Invite: no session for admin ${PRIMARY_OWNER_EMAIL}`);
          return;
        }
        const gToken = await getGraphToken(adminSession);
        if (!gToken) {
          log(`Invite: no Graph token for ${PRIMARY_OWNER_EMAIL}`);
          return;
        }
        const ok = await sendInviteEmail(normalizedEmail, gToken);
        log(`Invite: email ${ok ? "sent" : "FAILED"} to ${normalizedEmail}`);
      })();

      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: String(e) });
    }
  });

  app.delete("/org-admin/api/users/:email", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info || info.tier === "USER") {
      res.status(403).json({ error: "Forbidden" });
      return;
    }
    const email = String(req.params.email);
    const target = await getLcUserByEmail(email);
    const targetTier = getTier(email, target?.role ?? "USER", target?.orgRole);
    if (targetTier === "OWNER") {
      res.status(403).json({ error: "Cannot remove OWNER" });
      return;
    }
    try {
      await deleteLcUser(email);
      const st = await getSessionTokenByEmail(email);
      if (st) await removeUserToken(st);
      log(`${info.tier} ${info.email} removed ${email}`);
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: String(e) });
    }
  });

  app.post("/org-admin/api/users/:email/revoke-m365", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info || info.tier === "USER") {
      res.status(403).json({ error: "Forbidden" });
      return;
    }
    const email = String(req.params.email);
    const target = await getLcUserByEmail(email);
    const targetTier = getTier(email, target?.role ?? "USER", target?.orgRole);
    if (info.tier === "ADMIN" && targetTier !== "USER") {
      res.status(403).json({ error: "Admins can only revoke Users" });
      return;
    }
    try {
      const st = await getSessionTokenByEmail(email);
      if (st) await removeUserToken(st);
      log(`${info.tier} ${info.email} revoked M365 for ${email}`);
      res.json({ success: true, hadToken: !!st });
    } catch (e) {
      res.status(500).json({ error: String(e) });
    }
  });

  app.patch("/org-admin/api/users/:email/role", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info || info.tier === "USER") {
      res.status(403).json({ error: "Forbidden" });
      return;
    }
    const { role } = req.body ?? {};
    if (!role || !["ADMIN", "USER"].includes(role)) {
      res.status(400).json({ error: "role must be ADMIN or USER" });
      return;
    }
    const targetEmail = String(req.params.email);
    const target = await getLcUserByEmail(targetEmail);
    const targetTier = getTier(targetEmail, target?.role ?? "USER", target?.orgRole);
    if (targetTier === "OWNER") {
      res.status(403).json({ error: "Cannot change OWNER's role" });
      return;
    }
    try {
      await setLcUserOrgRole(targetEmail, role);
      log(`${info.tier} ${info.email} set ${targetEmail} → ${role}`);
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: String(e) });
    }
  });

  // ── Email-based forgot password flow ─────────────────────────────────────────
  // (The unauthenticated raw POST /org-admin/api/reset-password endpoint that
  // used to sit here — taking {email, newPassword} with zero proof of email
  // ownership — has been removed. It was confirmed dead code: nothing in
  // patch.js or elsewhere in this repo ever called it, and the flow below is
  // the real one: forgot-password emails a random 32-byte token, and only
  // reset-password-confirm, which verifies that token, may change a
  // password.)
  app.post("/org-admin/api/forgot-password", async (req: Request, res: Response) => {
    const { email } = req.body ?? {};
    if (!email || String(email).indexOf("@") === -1) {
      res.status(400).json({ error: "Valid email required" });
      return;
    }
    // Always respond success — don't reveal if email exists
    res.json({ success: true });

    void (async () => {
      const normalizedEmail = String(email).toLowerCase().trim();
      const user = await getLcUserByEmail(normalizedEmail);
      if (!user) {
        log(`ForgotPw: no user for ${normalizedEmail}`);
        return;
      }
      const token = crypto.randomBytes(32).toString("hex");
      const expiresAt = new Date(Date.now() + 60 * 60 * 1000);
      const db = await getDb();
      if (!db) return;
      await db.collection("passwordResetTokens").deleteMany({ email: normalizedEmail });
      await db.collection("passwordResetTokens").insertOne({
        email: normalizedEmail, token, expiresAt, used: false, createdAt: new Date(),
      });

      const adminSession = await getSessionTokenByEmail(PRIMARY_OWNER_EMAIL);
      if (!adminSession) {
        log(`ForgotPw: no session for admin ${PRIMARY_OWNER_EMAIL}`);
        return;
      }
      const gToken = await getGraphToken(adminSession);
      if (!gToken) {
        log(`ForgotPw: no Graph token for ${PRIMARY_OWNER_EMAIL}`);
        return;
      }
      const resetUrl = `${SERVER_BASE_URL}/reset-password?token=${token}`;
      const ok = await sendResetEmail(normalizedEmail, resetUrl, gToken);
      log(`ForgotPw: email ${ok ? "sent" : "FAILED"} to ${normalizedEmail}`);
    })();
  });

  app.get("/reset-password", async (req: Request, res: Response) => {
    const token = typeof req.query.token === "string" ? req.query.token : "";
    if (!token) {
      res.status(400).send(buildResetErrorPage("Invalid reset link."));
      return;
    }
    const db = await getDb();
    if (!db) {
      res.status(500).send(buildResetErrorPage("Server error."));
      return;
    }
    const record = await db.collection("passwordResetTokens").findOne({ token, used: false });
    if (!record || new Date() > (record.expiresAt as Date)) {
      res.status(400).send(buildResetErrorPage("This reset link has expired or already been used."));
      return;
    }
    res.send(buildResetPage(token));
  });

  app.post("/org-admin/api/reset-password-confirm", async (req: Request, res: Response) => {
    const { token, newPassword } = req.body ?? {};
    if (!token || !newPassword || String(newPassword).length < 8) {
      res.status(400).json({ error: "Invalid request" });
      return;
    }
    const db = await getDb();
    if (!db) {
      res.status(500).json({ error: "Server error" });
      return;
    }
    const record = await db.collection("passwordResetTokens").findOne({ token, used: false });
    if (!record || new Date() > (record.expiresAt as Date)) {
      res.status(400).json({ error: "Reset link expired or already used." });
      return;
    }
    const hash = await bcrypt.hash(newPassword, 10);
    await db.collection("users").updateOne({ email: record.email }, { $set: { password: hash, updatedAt: new Date() } });
    await db.collection("passwordResetTokens").updateOne({ token }, { $set: { used: true } });
    log(`Password reset confirmed for ${record.email}`);
    res.json({ success: true });
  });

  log("Org admin endpoints registered at /org-admin");
}
