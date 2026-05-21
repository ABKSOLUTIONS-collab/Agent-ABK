import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { AppConfig, ResolvedServer } from "../config/types";

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

// Client TTL: recreate connections every 8 minutes to stay well within
// any server-side timeout. Tokens are refreshed separately via Table Storage.
const CLIENT_TTL_MS = 8 * 60 * 1000;

interface ClientEntry {
  client: Client;
  createdAt: number;
}

/**
 * Forwards tool calls to Agent 365 MCP servers using a per-user bearer token.
 *
 * Key design decisions:
 * - No TokenProvider dependency — token is passed directly per request
 * - Per-user client cache keyed by (userToken + serverName) — full isolation
 * - Auto-reconnects on error or TTL expiry — never gets stuck
 * - Each user only calls tools on their own M365 tenant
 */
export class ToolForwarder {
  private config: AppConfig;
  private userToken: string;
  private servers: ResolvedServer[];

  // Cache: "serverName:tokenPrefix" -> ClientEntry
  private clientCache = new Map<string, ClientEntry>();

  constructor(config: AppConfig, userToken: string, servers: ResolvedServer[]) {
    this.config = config;
    this.userToken = userToken;
    this.servers = servers;
  }

  async callTool(
    toolName: string,
    args: Record<string, unknown>,
    targetServer: ResolvedServer
  ): Promise<{ content: Array<{ type: string; text: string }> }> {
    try {
      const client = await this.getOrCreateClient(targetServer);
      const result = await client.callTool({ name: toolName, arguments: args });
      return {
        content: (result.content as Array<{ type: string; text: string }>) ?? [
          { type: "text", text: JSON.stringify(result) },
        ],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      log(`Tool call failed for ${toolName} on ${targetServer.config.mcpServerName}: ${message}`);

      // On any error, invalidate the cached client and retry once with a fresh connection
      this.invalidateClient(targetServer);
      try {
        const freshClient = await this.getOrCreateClient(targetServer);
        const result = await freshClient.callTool({ name: toolName, arguments: args });
        return {
          content: (result.content as Array<{ type: string; text: string }>) ?? [
            { type: "text", text: JSON.stringify(result) },
          ],
        };
      } catch (retryErr) {
        const retryMessage = retryErr instanceof Error ? retryErr.message : String(retryErr);
        log(`Retry also failed for ${toolName}: ${retryMessage}`);
        return {
          content: [{
            type: "text",
            text: `Error calling tool "${toolName}": ${retryMessage}`,
          }],
        };
      }
    }
  }

  private cacheKey(server: ResolvedServer): string {
    // Use first 8 chars of token as key suffix — enough for isolation, not sensitive
    return `${server.config.mcpServerName}:${this.userToken.substring(0, 8)}`;
  }

  private invalidateClient(server: ResolvedServer): void {
    const key = this.cacheKey(server);
    const entry = this.clientCache.get(key);
    if (entry) {
      try { entry.client.close(); } catch { /* ignore */ }
      this.clientCache.delete(key);
    }
  }

  private async getOrCreateClient(server: ResolvedServer): Promise<Client> {
    const key = this.cacheKey(server);
    const now = Date.now();
    const entry = this.clientCache.get(key);

    // Return cached client if still within TTL
    if (entry && (now - entry.createdAt) < CLIENT_TTL_MS) {
      return entry.client;
    }

    // Close expired client if exists
    if (entry) {
      try { await entry.client.close(); } catch { /* ignore */ }
      this.clientCache.delete(key);
    }

    // Create fresh client with current user token
    const headers: Record<string, string> = {
      "User-Agent": "Agent365SDK/1.0.0 (ClaudeCodeBridge; Node.js)",
      Authorization: `Bearer ${this.userToken}`,
    };
    if (this.config.agenticAppId) {
      headers["x-ms-agentid"] = this.config.agenticAppId;
    }

    const transport = new StreamableHTTPClientTransport(
      new URL(server.url),
      { requestInit: { headers } }
    );

    const client = new Client({
      name: "agent365-claude-bridge",
      version: "1.0.0",
    });

    await client.connect(transport);
    this.clientCache.set(key, { client, createdAt: now });
    log(`Created new client for ${server.config.mcpServerName} (user: ${this.userToken.substring(0, 8)}...)`);
    return client;
  }

  async closeAll(): Promise<void> {
    for (const [, entry] of this.clientCache) {
      try { await entry.client.close(); } catch { /* ignore */ }
    }
    this.clientCache.clear();
  }
}