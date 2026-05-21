import { loadConfig } from "./config/configuration";
import { McpProxyServer } from "./proxy/mcp-proxy-server";

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

async function main(): Promise<void> {
  log("Starting Agent 365 Bridge...");

  const config = loadConfig();
  log(`Environment: ${config.nodeEnv}`);
  log(`Platform endpoint: ${config.mcpPlatformEndpoint}`);
  log(`Manifest servers: ${config.manifest.mcpServers.length}`);
  log("Auth: per-user OAuth tokens (no device code required)");

  // No TokenProvider, no server-level device code, no discovery callback.
  // Tools are discovered per-user on first request using their personal token.
  const proxy = new McpProxyServer(config);
  await proxy.start();

  const cleanup = async () => {
    log("Shutting down...");
    process.exit(0);
  };

  process.on("SIGINT", cleanup);
  process.on("SIGTERM", cleanup);
}

main().catch((err) => {
  log(`Fatal error: ${err}`);
  process.exit(1);
});