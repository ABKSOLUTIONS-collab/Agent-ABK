import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { AppConfig, DiscoveredTool, MCPServerConfig, ResolvedServer } from "../config/types";

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

/**
 * Discovers Agent 365 MCP servers and enumerates their tools.
 *
 * IMPORTANT: This class no longer uses TokenProvider or device code flow.
 * All discovery is done with a per-user bearer token passed directly.
 * This ensures:
 *  - No server-level auth that can expire
 *  - Each user only discovers tools they have access to
 *  - No manual intervention ever needed
 */
export class ServerDiscovery {
  private config: AppConfig;

  constructor(config: AppConfig) {
    this.config = config;
  }

  /**
   * Discovers all configured MCP servers using the provided user token.
   * Called once per user session, results are cached in McpProxyServer.
   */
  async discoverAll(userToken: string): Promise<ResolvedServer[]> {
    const serverConfigs = await this.getServerConfigs(userToken);
    const resolved: ResolvedServer[] = [];

    for (const serverConfig of serverConfigs) {
      try {
        const url = this.buildServerUrl(serverConfig);
        const tools = await this.discoverTools(serverConfig.mcpServerName, url, userToken);
        resolved.push({ config: serverConfig, url, tools });
        log(`Discovered ${tools.length} tools from ${serverConfig.mcpServerName}`);
      } catch (err) {
        log(`Failed to discover tools from ${serverConfig.mcpServerName}: ${err}`);
      }
    }

    return resolved;
  }

  private async getServerConfigs(userToken: string): Promise<MCPServerConfig[]> {
    if (this.config.agenticAppId) {
      return this.getFromGateway(userToken);
    }
    return this.getFromManifest();
  }

  private getFromManifest(): MCPServerConfig[] {
    const servers = this.config.manifest.mcpServers;
    if (servers.length === 0) {
      log("Warning: ToolingManifest.json contains no MCP servers");
    }
    return servers;
  }

  private async getFromGateway(userToken: string): Promise<MCPServerConfig[]> {
    const url = `${this.config.mcpPlatformEndpoint}/agents/${this.config.agenticAppId}/mcpServers`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${userToken}`,
        "Content-Type": "application/json",
      },
    });
    if (!response.ok) {
      throw new Error(`Gateway discovery failed: ${response.status} ${response.statusText}`);
    }
    const data = (await response.json()) as { mcpServers?: MCPServerConfig[] };
    return data.mcpServers ?? [];
  }

  private buildServerUrl(config: MCPServerConfig): string {
    if (config.url) return config.url;
    const base = this.config.mcpPlatformEndpoint.replace(/\/+$/, "");
    return `${base}/agents/servers/${config.mcpServerName}/`;
  }

  private async discoverTools(
    serverName: string,
    serverUrl: string,
    userToken: string
  ): Promise<DiscoveredTool[]> {
    const headers: Record<string, string> = {
      "User-Agent": "Agent365SDK/1.0.0 (ClaudeCodeBridge; Node.js)",
      Authorization: `Bearer ${userToken}`,
    };
    if (this.config.agenticAppId) {
      headers["x-ms-agentid"] = this.config.agenticAppId;
    }

    const transport = new StreamableHTTPClientTransport(new URL(serverUrl), {
      requestInit: { headers },
    });

    const client = new Client({ name: "agent365-bridge", version: "1.0.0" });
    await client.connect(transport);
    const result = await client.listTools();
    await client.close();

    return (result.tools ?? []).map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema as Record<string, unknown>,
      serverName,
    }));
  }
}