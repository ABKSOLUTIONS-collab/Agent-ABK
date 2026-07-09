import { TableClient, AzureNamedKeyCredential, TableEntity } from "@azure/data-tables";
import { createHash, randomUUID } from "node:crypto";

// ── Azure Table Storage config (same storage account as user-token-store.ts) ──
const ACCOUNT     = process.env.AZURE_STORAGE_ACCOUNT ?? "";
const ACCOUNT_KEY = process.env.AZURE_STORAGE_KEY ?? "";

const USERS_TABLE    = process.env.APP_USERS_TABLE          || "agent365appusers";
const RESETS_TABLE   = process.env.APP_PASSWORD_RESET_TABLE || "agent365apppasswordresets";
const SETTINGS_TABLE = process.env.APP_SETTINGS_TABLE       || "agent365appsettings";
const ERRORS_TABLE   = process.env.APP_ERROR_LOG_TABLE      || "agent365apperrorlog";

const DEFAULT_RESET_EMAIL_SENDER = "snikolaou@abk.gr";

function log(msg: string) {
  process.stderr.write(`[agent365-bridge] ${msg}\n`);
}

function getClient(tableName: string): TableClient {
  if (!ACCOUNT || !ACCOUNT_KEY) {
    throw new Error("AZURE_STORAGE_ACCOUNT and AZURE_STORAGE_KEY must be set");
  }
  const cred = new AzureNamedKeyCredential(ACCOUNT, ACCOUNT_KEY);
  return new TableClient(`https://${ACCOUNT}.table.core.windows.net`, tableName, cred);
}

async function ensureTable(tableName: string): Promise<void> {
  const client = getClient(tableName);
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err; // 409 = table already exists
  }
}

export async function ensureAppTables(): Promise<void> {
  await Promise.all([
    ensureTable(USERS_TABLE),
    ensureTable(RESETS_TABLE),
    ensureTable(SETTINGS_TABLE),
    ensureTable(ERRORS_TABLE),
  ]);
}

export type AppRole = "admin" | "user";

export interface AppUser {
  email: string;
  passwordHash: string;
  role: AppRole;
  createdAt: number;
  isActive: boolean;
}

interface UserEntity extends TableEntity, AppUser {}

function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

// ── Users ───────────────────────────────────────────────────────────────────

export async function getUser(email: string): Promise<AppUser | null> {
  try {
    const client = getClient(USERS_TABLE);
    const entity = await client.getEntity<UserEntity>("users", normalizeEmail(email));
    return {
      email: entity.email,
      passwordHash: entity.passwordHash,
      role: entity.role,
      createdAt: entity.createdAt,
      isActive: entity.isActive,
    };
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) log(`app-store getUser error: ${err}`);
    return null;
  }
}

export async function upsertUser(user: AppUser): Promise<void> {
  const client = getClient(USERS_TABLE);
  const entity: UserEntity = {
    partitionKey: "users",
    rowKey: normalizeEmail(user.email),
    email: normalizeEmail(user.email),
    passwordHash: user.passwordHash,
    role: user.role,
    createdAt: user.createdAt,
    isActive: user.isActive,
  };
  await client.upsertEntity(entity, "Replace");
}

export async function deleteUser(email: string): Promise<void> {
  const client = getClient(USERS_TABLE);
  await client.deleteEntity("users", normalizeEmail(email));
}

export async function listUsers(): Promise<AppUser[]> {
  const client = getClient(USERS_TABLE);
  const users: AppUser[] = [];
  for await (const entity of client.listEntities<UserEntity>({
    queryOptions: { filter: `PartitionKey eq 'users'` },
  })) {
    users.push({
      email: entity.email,
      passwordHash: entity.passwordHash,
      role: entity.role,
      createdAt: entity.createdAt,
      isActive: entity.isActive,
    });
  }
  return users;
}

export async function countUsers(): Promise<number> {
  const client = getClient(USERS_TABLE);
  let count = 0;
  for await (const _ of client.listEntities({ queryOptions: { filter: `PartitionKey eq 'users'` } })) {
    count++;
  }
  return count;
}

// ── Password reset tokens ────────────────────────────────────────────────────

interface ResetEntity extends TableEntity {
  email: string;
  expiresAt: number;
  used: boolean;
}

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}

export async function createResetToken(email: string, ttlMs: number): Promise<string> {
  const token = randomUUID() + randomUUID();
  const client = getClient(RESETS_TABLE);
  const entity: ResetEntity = {
    partitionKey: "pwreset",
    rowKey: hashToken(token),
    email: normalizeEmail(email),
    expiresAt: Date.now() + ttlMs,
    used: false,
  };
  await client.upsertEntity(entity, "Replace");
  return token;
}

export async function consumeResetToken(token: string): Promise<string | null> {
  const client = getClient(RESETS_TABLE);
  const rowKey = hashToken(token);
  try {
    const entity = await client.getEntity<ResetEntity>("pwreset", rowKey);
    if (entity.used || Date.now() > entity.expiresAt) return null;
    await client.upsertEntity<ResetEntity>({ ...entity, used: true }, "Replace");
    return entity.email;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) log(`app-store consumeResetToken error: ${err}`);
    return null;
  }
}

// ── App settings ──────────────────────────────────────────────────────────────

interface SettingsEntity extends TableEntity {
  resetEmailSender: string;
}

export async function getResetEmailSender(): Promise<string> {
  try {
    const client = getClient(SETTINGS_TABLE);
    const entity = await client.getEntity<SettingsEntity>("config", "general");
    return entity.resetEmailSender || DEFAULT_RESET_EMAIL_SENDER;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 404) log(`app-store getResetEmailSender error: ${err}`);
    return DEFAULT_RESET_EMAIL_SENDER;
  }
}

export async function setResetEmailSender(sender: string): Promise<void> {
  const client = getClient(SETTINGS_TABLE);
  const entity: SettingsEntity = {
    partitionKey: "config",
    rowKey: "general",
    resetEmailSender: sender.trim(),
  };
  await client.upsertEntity(entity, "Replace");
}

// ── Error / health log ────────────────────────────────────────────────────────

export interface AppErrorLogEntry {
  timestamp: number;
  source: string;
  message: string;
  detail?: string;
}

interface ErrorEntity extends TableEntity, AppErrorLogEntry {}

function invertedTimestampKey(): string {
  // Sorting rowKey ascending yields newest-first ordering.
  const inverted = (9999999999999 - Date.now()).toString().padStart(13, "0");
  return `${inverted}-${randomUUID()}`;
}

export async function logAppError(source: string, message: string, detail?: string): Promise<void> {
  try {
    const client = getClient(ERRORS_TABLE);
    const entity: ErrorEntity = {
      partitionKey: "errors",
      rowKey: invertedTimestampKey(),
      timestamp: Date.now(),
      source,
      message,
      detail: detail ?? "",
    };
    await client.upsertEntity(entity, "Replace");
  } catch (err) {
    // Never let logging failures break the request path.
    log(`app-store logAppError failed: ${err}`);
  }
}

export async function listAppErrors(limit = 100): Promise<AppErrorLogEntry[]> {
  const client = getClient(ERRORS_TABLE);
  const entries: AppErrorLogEntry[] = [];
  for await (const entity of client.listEntities<ErrorEntity>({
    queryOptions: { filter: `PartitionKey eq 'errors'` },
  })) {
    entries.push({
      timestamp: entity.timestamp,
      source: entity.source,
      message: entity.message,
      detail: entity.detail,
    });
    if (entries.length >= limit) break;
  }
  return entries;
}
