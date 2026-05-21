import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import {
    AuthenticationRecord,
    serializeAuthenticationRecord,
    deserializeAuthenticationRecord,
} from "@azure/identity";

/**
 * Manages persistent storage of the Azure Identity AuthenticationRecord.
 *
 * The AuthenticationRecord is a non-sensitive object that identifies the
 * authenticated user. When combined with @azure/identity-cache-persistence,
 * it allows the DeviceCodeCredential to acquire tokens silently without
 * requiring interactive sign-in on each launch.
 *
 * Storage location: ~/.agent365-bridge/auth-record.json
 */

// Allow overriding the cache directory via env var (same as tools-cache.ts).
const CACHE_DIR = process.env.AGENT365_CACHE_DIR ?? path.join(os.homedir(), ".agent365-bridge");
const AUTH_RECORD_FILE = path.join(CACHE_DIR, "auth-record.json");

function log(message: string): void {
    process.stderr.write(`[agent365-bridge] ${message}\n`);
}

/**
 * Saves an AuthenticationRecord to disk for reuse across sessions.
 */
export function saveAuthRecord(record: AuthenticationRecord): void {
    if (!fs.existsSync(CACHE_DIR)) {
        fs.mkdirSync(CACHE_DIR, { recursive: true });
    }

    const serialized = serializeAuthenticationRecord(record);
    fs.writeFileSync(AUTH_RECORD_FILE, serialized, "utf-8");
    log(`Authentication record saved to ${AUTH_RECORD_FILE}`);
}

/**
 * Loads a previously saved AuthenticationRecord from disk.
 * Returns null if no record exists or if it cannot be read.
 */
export function loadAuthRecord(): AuthenticationRecord | null {
    try {
        if (!fs.existsSync(AUTH_RECORD_FILE)) {
            return null;
        }

        const serialized = fs.readFileSync(AUTH_RECORD_FILE, "utf-8");
        const record = deserializeAuthenticationRecord(serialized);
        log("Loaded cached authentication record (silent auth enabled)");
        return record;
    } catch (err) {
        log(`Failed to load cached auth record: ${err}`);
        return null;
    }
}

/**
 * Deletes the cached AuthenticationRecord from disk.
 */
export function clearAuthRecord(): void {
    try {
        if (fs.existsSync(AUTH_RECORD_FILE)) {
            fs.unlinkSync(AUTH_RECORD_FILE);
            log(`Removed cached auth record from ${AUTH_RECORD_FILE}`);
        } else {
            log("No cached auth record found to remove.");
        }
    } catch (err) {
        log(`Failed to clear auth record: ${err}`);
    }
}

/** Returns the path to the auth record file (for display purposes). */
export function getAuthRecordPath(): string {
    return AUTH_RECORD_FILE;
}