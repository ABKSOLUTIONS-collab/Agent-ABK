import * as crypto from "crypto";
import { Express, Request, Response } from "express";
import { MongoClient, Db, ObjectId } from "mongodb";
import bcrypt from "bcryptjs";
import { TableClient, AzureNamedKeyCredential } from "@azure/data-tables";
import { getSessionTokenByEmail, removeUserToken, getGraphToken } from "../auth/user-token-store";

const SERVER_BASE_URL = (process.env.SERVER_BASE_URL ?? "").replace(/\/$/, "");

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

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    const payload = Buffer.from(parts[1], "base64url").toString("utf8");
    return JSON.parse(payload);
  } catch {
    return null;
  }
}

async function resolveByLcToken(lcToken: string): Promise<AuthInfo | null> {
  const payload = decodeJwtPayload(lcToken);
  if (!payload) return null;
  const jwtEmail = payload.email as string | undefined;
  const jwtId = payload.id as string | undefined;
  if (!jwtEmail && !jwtId) return null;

  const db = await getDb();
  if (!db) return null;
  try {
    let user: LcUser | null = null;
    if (jwtEmail) {
      user = (await db.collection<LcUser>("users").findOne({ email: jwtEmail })) ?? null;
    }
    if (!user && jwtId) {
      user = (await db.collection<LcUser>("users").findOne({ _id: new ObjectId(jwtId) })) ?? null;
    }
    if (!user) return null;
    return { email: user.email, tier: getTier(user.email, user.role, user.orgRole) };
  } catch (e) {
    log(`resolveByLcToken error: ${e}`);
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

async function resolveByLcEmail(email: string): Promise<AuthInfo | null> {
  if (!email || !email.includes("@")) return null;
  const user = await getLcUserByEmail(email);
  if (!user) return null;
  return { email: user.email, tier: getTier(user.email, user.role, user.orgRole) };
}

async function resolveAuth(req: Request): Promise<AuthInfo | null> {
  const lcToken = qstr(req.query.lc_token);
  const lcEmail = qstr(req.query.lc_email);
  const m365Token = qstr(req.query.token);
  if (lcToken) return resolveByLcToken(lcToken);
  if (lcEmail) return resolveByLcEmail(lcEmail);
  if (m365Token) return resolveByM365Token(m365Token);
  return null;
}

// ── HTML ─────────────────────────────────────────────────────────────────────

function buildHtml(email: string, tier: Tier, embed: boolean): string {
  const authParamJs = `
const _p = new URLSearchParams(location.search);
const LC_TOKEN   = _p.get('lc_token') || '';
const LC_EMAIL   = _p.get('lc_email') || '';
const M365_TOKEN = _p.get('token')    || '';
function authParam() {
  if (LC_TOKEN)   return 'lc_token=' + encodeURIComponent(LC_TOKEN);
  if (LC_EMAIL)   return 'lc_email=' + encodeURIComponent(LC_EMAIL);
  return 'token=' + encodeURIComponent(M365_TOKEN);
}`;

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
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Organization Settings</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    @keyframes abkFadeUp{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes abkPulse{0%,100%{opacity:1}50%{opacity:.45}}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#fafafa;color:#1a1a1a;min-height:100vh;font-size:14px;-webkit-font-smoothing:antialiased}
    a{text-decoration:none;color:inherit}
    .layout{display:flex;min-height:100vh;animation:abkFadeUp .25s ease}
    /* Sidebar */
    .sidebar{width:192px;flex-shrink:0;padding:16px 8px;border-right:1px solid #ececec;background:#fff}
    .sb-section{margin-bottom:16px}
    .sb-label{font-size:11px;font-weight:500;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;padding:0 8px;margin-bottom:2px}
    .sb-item{display:flex;align-items:center;padding:7px 8px;border-radius:7px;font-size:13px;color:#374151;cursor:pointer;gap:8px;margin-bottom:1px;transition:background .15s ease,color .15s ease}
    .sb-item.active{background:#eaf2fe;color:#0066cc;font-weight:500}
    .sb-item:hover:not(.active){background:#f9fafb}
    /* Main */
    .main{flex:1;padding:${embed ? "20px 24px" : "32px 40px"};max-width:900px}
    h1{font-size:18px;font-weight:600;color:#111;margin-bottom:20px;letter-spacing:-.01em}
    /* Overview */
    .overview{display:flex;background:#fff;border:1px solid #ececec;border-radius:12px;margin-bottom:24px;overflow:hidden;box-shadow:0 1px 2px rgba(16,24,40,.04)}
    .ov-item{flex:1;padding:16px 20px;border-right:1px solid #f0f0f0;transition:background .15s ease}
    .ov-item:last-child{border-right:none}
    .ov-label{font-size:11px;color:#9ca3af;margin-bottom:5px;font-weight:500;text-transform:uppercase;letter-spacing:.05em}
    .ov-val{font-size:14px;font-weight:600;color:#111}
    .ov-sub{font-size:12px;color:#0066cc;margin-top:3px;font-weight:500}
    /* Members header */
    .m-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
    .m-title{font-size:15px;font-weight:600;color:#111}
    .btn-primary{background:#0066cc;color:#fff;border:none;border-radius:8px;padding:7px 15px;font-size:13px;font-weight:500;cursor:pointer;transition:background .15s ease,transform .1s ease,box-shadow .15s ease;box-shadow:0 1px 2px rgba(0,102,204,.25)}
    .btn-primary:hover{background:#0055b3;box-shadow:0 2px 8px rgba(0,102,204,.3)}
    .btn-primary:active{transform:translateY(1px)}
    /* Table */
    .table-wrap{background:#fff;border:1px solid #ececec;border-radius:12px;overflow:hidden;box-shadow:0 1px 2px rgba(16,24,40,.04)}
    .table-toolbar{display:flex;align-items:center;gap:10px;padding:10px 16px;border-bottom:1px solid #f3f4f6}
    .search-wrap{position:relative;display:flex;align-items:center}
    .search-wrap svg{position:absolute;left:9px;pointer-events:none;color:#b0b3b9}
    .search{border:1px solid #e5e7eb;border-radius:7px;padding:6px 11px 6px 28px;font-size:13px;outline:none;width:200px;color:#111;background:#fff;transition:border-color .15s ease,box-shadow .15s ease}
    .search:focus{border-color:#0066cc;box-shadow:0 0 0 3px rgba(0,102,204,.12)}
    .count-lbl{font-size:13px;color:#9ca3af}
    table{width:100%;border-collapse:collapse}
    th{padding:10px 16px;text-align:left;font-size:11px;font-weight:500;color:#9ca3af;border-bottom:1px solid #ececec;text-transform:uppercase;letter-spacing:.05em}
    td{padding:12px 16px;font-size:13px;border-bottom:1px solid #f5f5f5;vertical-align:middle}
    tr:last-child td{border-bottom:none}
    tr{transition:background .12s ease}
    tr:hover td{background:#fafcff}
    /* Name cell */
    .name-cell{display:flex;align-items:center;gap:10px}
    .avatar{width:30px;height:30px;border-radius:50%;color:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;flex-shrink:0;text-transform:uppercase;box-shadow:inset 0 0 0 1px rgba(255,255,255,.15)}
    .name-primary{font-size:13px;font-weight:500;color:#111}
    .name-secondary{font-size:12px;color:#9ca3af}
    .you-badge{font-size:11px;color:#0066cc;background:#eaf2fe;border-radius:5px;padding:1px 6px;margin-left:5px;font-weight:500}
    /* Role */
    .role-text{font-size:13px;color:#374151}
    .role-select{padding:4px 28px 4px 8px;border:1px solid #e5e7eb;border-radius:7px;font-size:13px;cursor:pointer;background:#fff;color:#374151;outline:none;appearance:auto;transition:border-color .15s ease}
    .role-select:hover{border-color:#c9ccd1}
    .role-select:focus{border-color:#0066cc;box-shadow:0 0 0 3px rgba(0,102,204,.12)}
    /* Actions */
    .btn-remove{background:none;border:none;cursor:pointer;padding:4px 10px;border-radius:6px;font-size:12.5px;color:#6b7280;white-space:nowrap;transition:background .15s ease,color .15s ease}
    .btn-remove:hover{background:#fef2f2;color:#dc2626}
    /* Error */
    .err-bar{background:#fef2f2;border:1px solid #fca5a5;border-radius:9px;padding:10px 14px;color:#dc2626;font-size:13px;margin-bottom:16px;display:none;animation:abkFadeUp .2s ease}
    .empty-cell{text-align:center;padding:56px 20px;color:#9ca3af;font-size:13px}
    .empty-cell svg{display:block;margin:0 auto 12px;color:#d1d5db}
    .skeleton{background:linear-gradient(90deg,#f0f0f0,#f6f6f6,#f0f0f0);border-radius:6px;animation:abkPulse 1.3s ease infinite}
    /* Add member modal */
    .add-input{width:100%;padding:9px 12px;border:1px solid #e5e7eb;border-radius:8px;font-size:13px;margin-bottom:10px;outline:none;box-sizing:border-box;transition:border-color .15s ease,box-shadow .15s ease}
    .add-input:focus{border-color:#0066cc;box-shadow:0 0 0 3px rgba(0,102,204,.12)}
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
          ${tier !== "USER" ? "<th></th>" : ""}
        </tr></thead>
        <tbody id="tbody">${[0, 1, 2].map(() => `<tr><td><div class="name-cell"><div class="skeleton" style="width:30px;height:30px;border-radius:50%"></div><div style="flex:1"><div class="skeleton" style="width:120px;height:11px;margin-bottom:6px"></div><div class="skeleton" style="width:160px;height:10px"></div></div></div></td><td><div class="skeleton" style="width:60px;height:11px"></div></td>${tier !== "USER" ? "<td></td>" : ""}</tr>`).join("")}</tbody>
      </table>
    </div>
  </div>
</div>
<!-- Add member modal -->
<div id="addModal" style="display:none;position:fixed;inset:0;z-index:9999;background:rgba(17,17,17,0.5);align-items:center;justify-content:center;backdrop-filter:blur(2px);">
  <div style="background:#fff;border-radius:14px;padding:28px 32px;width:380px;max-width:95vw;box-shadow:0 24px 60px rgba(0,0,0,.22);animation:abkFadeUp .18s ease;">
    <h2 style="font-size:15px;font-weight:600;color:#111;margin-bottom:6px">Add member</h2>
    <p style="font-size:13px;color:#6b7280;margin-bottom:18px">Create a new @abk.gr account.</p>
    <div id="addMsg"></div>
    <input id="addName"  class="add-input" type="text"     placeholder="Full name">
    <input id="addEmail" class="add-input" type="email"    placeholder="you@abk.gr">
    <input id="addPw1"   class="add-input" type="password" placeholder="Password (min 8 chars)">
    <input id="addPw2"   class="add-input" type="password" placeholder="Confirm password" style="margin-bottom:16px">
    <button id="addSubmitBtn" class="btn-primary" onclick="submitAdd()" style="width:100%;padding:10px;margin-bottom:8px">Create account</button>
    <button onclick="hideAddModal()" style="width:100%;padding:8px;background:none;border:none;font-size:13px;color:#6b7280;cursor:pointer;border-radius:8px;transition:background .15s ease" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='none'">Cancel</button>
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
  if (!users.length) { tb.innerHTML = '<tr><td class="empty-cell" colspan="3"><svg width="32" height="32" viewBox="0 0 16 16" fill="currentColor"><path d="M15 14s1 0 1-1-1-4-5-4-5 3-5 4 1 1 1 1h8zm-9.978-1A.261.261 0 015 13c0-.366.268-1.14 1.004-1.844C6.717 10.48 7.742 10 9 10c.37 0 .724.043 1.05.12C9.5 10.48 9 11.25 9 12c0 .273.04.54.097.8H5.022zM4.5 8a2.5 2.5 0 100-5 2.5 2.5 0 000 5z"/></svg>No members found</td></tr>'; return; }
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

    return \`<tr>
      <td><div class="name-cell">
        <div class="avatar" style="background:\${color}">\${initials(name)}</div>
        <div><div class="name-primary">\${name}\${youBadge}</div><div class="name-secondary">\${u.email}</div></div>
      </div></td>
      <td>\${roleCell}</td>
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
  ['addName','addEmail','addPw1','addPw2'].forEach(function(id){ document.getElementById(id).value=''; });
  var msg=document.getElementById('addMsg'); msg.style.display='none';
  var btn=document.getElementById('addSubmitBtn'); btn.disabled=false; btn.textContent='Create account'; btn.style.display='';
  var m=document.getElementById('addModal'); m.style.display='flex';
}
function hideAddModal() { document.getElementById('addModal').style.display='none'; }
document.getElementById('addModal').addEventListener('click', function(e){ if(e.target===this) hideAddModal(); });

async function submitAdd() {
  var name=document.getElementById('addName').value.trim();
  var email=document.getElementById('addEmail').value.trim().toLowerCase();
  var pw1=document.getElementById('addPw1').value;
  var pw2=document.getElementById('addPw2').value;
  var msg=document.getElementById('addMsg');
  function showAddMsg(txt,isErr){
    msg.style.display='block';
    msg.style.background=isErr?'#fef2f2':'#f0fdf4';
    msg.style.color=isErr?'#dc2626':'#16a34a';
    msg.textContent=txt;
  }
  if(!name){showAddMsg('Enter a full name.',true);return;}
  if(!email||email.indexOf('@')===-1){showAddMsg('Enter a valid email.',true);return;}
  if(pw1.length<8){showAddMsg('Password must be at least 8 characters.',true);return;}
  if(pw1!==pw2){showAddMsg('Passwords do not match.',true);return;}
  var btn=document.getElementById('addSubmitBtn');
  btn.disabled=true; btn.textContent='Creating…';
  try {
    const r=await fetch('/org-admin/api/users?'+authParam(),{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({name,email,password:pw1})
    });
    const d=await r.json();
    if(r.ok){
      showAddMsg('Account created!',false);
      btn.style.display='none';
      setTimeout(function(){ hideAddModal(); load(); },1200);
    } else {
      showAddMsg(d.error||'Error creating account.',true);
      btn.disabled=false; btn.textContent='Create account';
    }
  } catch(e){ showAddMsg('Connection error.',true); btn.disabled=false; btn.textContent='Create account'; }
}

load();
</script>
</body>
</html>`;
}

// ── Email prompt HTML (shown in embed mode when no auth provided) ─────────────

function buildEmailPromptHtml(showError: boolean): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Organization Settings</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    @keyframes abkFadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#fafafa;min-height:100vh;display:flex;align-items:center;justify-content:center}
    .card{background:#fff;border:1px solid #f0f0f0;border-radius:14px;padding:36px 40px;box-shadow:0 8px 28px rgba(16,24,40,.08);width:300px;animation:abkFadeUp .25s ease}
    h2{font-size:16px;font-weight:600;color:#111;margin-bottom:8px}
    p{color:#777;font-size:13px;margin-bottom:20px;line-height:1.5}
    input{width:100%;padding:10px 14px;border:1px solid #e3e3e3;border-radius:9px;font-size:14px;margin-bottom:12px;outline:none;transition:border-color .15s ease,box-shadow .15s ease}
    input:focus{border-color:#0066cc;box-shadow:0 0 0 3px rgba(0,102,204,.14)}
    button{width:100%;padding:11px;background:#0066cc;color:#fff;border:none;border-radius:9px;font-size:14px;font-weight:500;cursor:pointer;transition:background .15s ease,box-shadow .15s ease}
    button:hover{background:#0055bb;box-shadow:0 2px 10px rgba(0,102,204,.28)}
    .err{color:#dc2626;font-size:12px;margin-bottom:12px;background:#fef2f2;padding:8px 10px;border-radius:7px}
  </style>
</head>
<body>
  <div class="card">
    <h2>Organization Settings</h2>
    <p>Enter your ABK email to continue</p>
    ${showError ? '<div class="err">Email not recognized. Use your @abk.gr account.</div>' : ""}
    <input type="email" id="emailInput" placeholder="you@abk.gr" autofocus>
    <button onclick="go()">Continue</button>
  </div>
  <script>
    var saved = localStorage.getItem('abk_admin_email');
    if (saved) document.getElementById('emailInput').value = saved;
    document.getElementById('emailInput').addEventListener('keydown', function(e){ if(e.key==='Enter') go(); });
    function go() {
      var e = document.getElementById('emailInput').value.trim();
      if (!e || !e.includes('@')) return;
      localStorage.setItem('abk_admin_email', e);
      // Tell the parent LibreChat page to cache this email in its own localStorage
      try { window.parent.postMessage({ type: 'abk_admin_email', email: e }, '*'); } catch(ex) {}
      var p = new URLSearchParams(location.search);
      p.set('lc_email', e);
      location.href = '/org-admin?' + p.toString();
    }
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
  input:focus{border-color:#0066CC;box-shadow:0 0 0 3px rgba(0,102,204,.14);}
  button{width:100%;padding:11px;background:#0066CC;color:#fff;border:none;border-radius:9px;font-size:14px;font-weight:500;cursor:pointer;transition:background .15s ease,box-shadow .15s ease}
  button:hover{background:#0055b3;box-shadow:0 2px 10px rgba(0,102,204,.28);}
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
      document.querySelector('.card').innerHTML='<img src="${SERVER_BASE_URL}/assets/abk-logo.png" alt="ABK" style="height:36px;margin-bottom:20px;display:block;"><h2 style="color:#0066CC">Password Updated!</h2><p>Your password has been changed. You can now <a href="${SERVER_BASE_URL.replace("agent365-bridge", "abkagent-backup").replace(/agent365.*/, "")}" style="color:#0066CC">log in</a>.</p>';
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
                <h2 style="color:#0066CC;font-size:20px;margin:0 0 12px;">Reset your password</h2>
                <p style="color:#444;line-height:1.6;">We received a request to reset your <strong>ABK Assistant</strong> password. Click below to set a new password. This link expires in <strong>1 hour</strong>.</p>
                <a href="${resetUrl}" style="display:inline-block;margin:24px 0;padding:12px 28px;background:#0066CC;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;">Reset Password</a>
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

// ── Express endpoints ──────────────────────────────────────────────────────────

export function registerOrgAdminEndpoints(app: Express): void {
  app.get("/org-admin", async (req: Request, res: Response) => {
    res.removeHeader("X-Frame-Options");
    const embed = qstr(req.query.embed) === "1";
    const info = await resolveAuth(req);
    if (!info) {
      if (embed) {
        const hadAttempt = !!(qstr(req.query.lc_email) || qstr(req.query.lc_token) || qstr(req.query.token));
        res.setHeader("Content-Type", "text/html");
        res.status(401).send(buildEmailPromptHtml(hadAttempt));
      } else {
        res.status(401).send("Unauthorized. Open via ABK Assistant.");
      }
      return;
    }
    res.setHeader("Content-Type", "text/html");
    res.send(buildHtml(info.email, info.tier, embed));
  });

  app.get("/org-admin/api/users", async (req: Request, res: Response) => {
    const info = await resolveAuth(req);
    if (!info) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    try {
      const lcUsers = await getLcUsers();
      const users = lcUsers.map((u) => ({
        email: u.email,
        name: u.name || u.username || u.email.split("@")[0],
        tier: getTier(u.email, u.role, u.orgRole),
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
    const { name, email, password } = req.body ?? {};
    if (!name || !email || !password) {
      res.status(400).json({ error: "name, email, and password are required" });
      return;
    }
    const normalizedEmail = String(email).toLowerCase().trim();
    if (String(password).length < 8) {
      res.status(400).json({ error: "Password must be at least 8 characters" });
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
      const hash = await bcrypt.hash(password, 10);
      const username = normalizedEmail.split("@")[0].replace(/[^a-zA-Z0-9_-]/g, "").toLowerCase();
      await db.collection("users").insertOne({
        name,
        username,
        email: normalizedEmail,
        password: hash,
        role: "USER",
        orgRole: "USER",
        provider: "local",
        emailVerified: true,
        createdAt: new Date(),
        updatedAt: new Date(),
        __v: 0,
      });
      log(`${info.tier} ${info.email} created account for ${normalizedEmail}`);
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

  // Password reset — no auth required (user is locked out).
  app.options("/org-admin/api/reset-password", (_req: Request, res: Response) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    res.sendStatus(204);
  });

  app.post("/org-admin/api/reset-password", async (req: Request, res: Response) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    const { email, newPassword } = req.body ?? {};
    if (!email || !newPassword) {
      res.status(400).json({ error: "email and newPassword required" });
      return;
    }
    if (String(newPassword).length < 8) {
      res.status(400).json({ error: "Password must be at least 8 characters" });
      return;
    }
    const user = await getLcUserByEmail(email);
    if (!user) {
      res.status(404).json({ error: "No account found with that email" });
      return;
    }
    try {
      const hash = await bcrypt.hash(newPassword, 10);
      const db = await getDb();
      await db!.collection("users").updateOne({ email }, { $set: { password: hash } });
      log(`Password reset for ${email}`);
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: String(e) });
    }
  });

  // ── Email-based forgot password flow ─────────────────────────────────────────
  app.post("/org-admin/api/forgot-password", async (req: Request, res: Response) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
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
    res.setHeader("Access-Control-Allow-Origin", "*");
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

  // Lightweight endpoint: returns just the current user's tier (for patch.js auto-fetch)
  app.get("/org-admin/api/my-tier", async (req: Request, res: Response) => {
    res.removeHeader("X-Frame-Options");
    res.setHeader("Access-Control-Allow-Origin", "*");
    const info = await resolveAuth(req);
    if (!info) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    res.json({ tier: info.tier });
  });

  log("Org admin endpoints registered at /org-admin");
}
