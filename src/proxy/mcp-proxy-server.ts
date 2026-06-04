import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { ResolvedServer } from "../config/types";
import { ToolForwarder } from "./tool-forwarder";
import { CachedTool, loadToolsCache, saveToolsCache } from "../auth/tools-cache";
import { getUserToken, getGraphToken, getTokenEmail, pruneExpiredTokens } from "../auth/user-token-store";
import { registerOAuthEndpoints } from "../auth/oauth-handler";
import { SHAREPOINT_TOOLS, SHAREPOINT_TOOL_NAMES, SharePointToolHandler } from "../tools/sharepoint-tools";
import { OCR_TOOL, OCR_TOOL_NAMES, OcrToolHandler } from "../tools/ocr-tool";
import {
  SIGNATURE_STYLE_TOOL,
  SIGNATURE_STYLE_TOOL_NAME,
  SET_SIGNATURE_TOOL,
  SET_SIGNATURE_TOOL_NAME,
  EMAIL_TOOLS_REQUIRING_SIGNATURE,
  handleGetSignatureStyle,
  handleSetSignature,
} from "../tools/signature-style-tool";
import { getUserSignature, appendSignature } from "../tools/signature-service";
import { ServerDiscovery } from "../discovery/server-discovery";
import { AppConfig } from "../config/types";
import express from "express";

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

interface ToolRegistryEntry {
  uniqueName: string;
  originalName: string;
  serverName: string;
  toolDef: Tool;
}

interface UserDiscoveryCache {
  servers: ResolvedServer[];
  toolRegistry: Map<string, ToolRegistryEntry>;
  forwarder: ToolForwarder;
  discoveredAt: number;
}

const DISCOVERY_TTL_MS = 25 * 60 * 1000;

export class McpProxyServer {
  private config: AppConfig;
  private sharedToolRegistry = new Map<string, ToolRegistryEntry>();
  private userCache = new Map<string, UserDiscoveryCache>();
  private discoveryInFlight = new Map<string, Promise<UserDiscoveryCache>>();

  constructor(config: AppConfig) {
    this.config = config;
    pruneExpiredTokens();
    // Load tools cache async at startup — non-blocking
    void this.initToolsCache();
  }

  private async initToolsCache(): Promise<void> {
    const cached = await loadToolsCache();
    if (cached && cached.length > 0) {
      this.rebuildSharedRegistry(cached);
      log(`Loaded ${this.sharedToolRegistry.size} tools from Table Storage cache`);
    }
  }

  private rebuildSharedRegistry(tools: CachedTool[]) {
    this.sharedToolRegistry.clear();
    const counts = new Map<string, number>();
    for (const t of tools) counts.set(t.name, (counts.get(t.name) || 0) + 1);
    for (const t of tools) {
      const uniqueName = (counts.get(t.name) || 0) > 1 ? `${t.name}_${t.serverName}` : t.name;
      this.sharedToolRegistry.set(uniqueName, {
        uniqueName,
        originalName: t.name,
        serverName: t.serverName,
        toolDef: {
          name: uniqueName,
          description: t.description,
          inputSchema: sanitizeSchema(t.inputSchema) as Tool["inputSchema"],
        },
      });
    }
  }

  private buildRegistry(servers: ResolvedServer[]): Map<string, ToolRegistryEntry> {
    const registry = new Map<string, ToolRegistryEntry>();
    const counts = new Map<string, number>();
    for (const s of servers) for (const t of s.tools) counts.set(t.name, (counts.get(t.name) || 0) + 1);
    for (const s of servers) {
      for (const t of s.tools) {
        const uniqueName = (counts.get(t.name) || 0) > 1 ? `${t.name}_${s.config.mcpServerName}` : t.name;
        registry.set(uniqueName, {
          uniqueName,
          originalName: t.name,
          serverName: s.config.mcpServerName,
          toolDef: {
            name: uniqueName,
            description: t.description,
            inputSchema: sanitizeSchema(t.inputSchema as Record<string, unknown>) as Tool["inputSchema"],
          },
        });
      }
    }
    return registry;
  }

  private async getUserCache(sessionToken: string, userToken: string): Promise<UserDiscoveryCache | null> {
    const cached = this.userCache.get(sessionToken);
    const now = Date.now();

    if (cached && (now - cached.discoveredAt) < DISCOVERY_TTL_MS) return cached;

    const inFlight = this.discoveryInFlight.get(sessionToken);
    if (inFlight) return inFlight;

    const promise = this.runDiscovery(sessionToken, userToken);
    this.discoveryInFlight.set(sessionToken, promise);
    promise.finally(() => this.discoveryInFlight.delete(sessionToken));

    if (cached) {
      log(`Refreshing tools in background for ${sessionToken.substring(0, 8)}...`);
      return cached;
    }

    try {
      return await promise;
    } catch (err) {
      log(`Discovery failed for ${sessionToken.substring(0, 8)}...: ${err}`);
      return null;
    }
  }

  private async runDiscovery(sessionToken: string, userToken: string): Promise<UserDiscoveryCache> {
    log(`Discovering tools for user ${sessionToken.substring(0, 8)}...`);
    const discovery = new ServerDiscovery(this.config);
    const servers = await discovery.discoverAll(userToken);
    const toolRegistry = this.buildRegistry(servers);
    const forwarder = new ToolForwarder(this.config, userToken, servers);

    const entry: UserDiscoveryCache = { servers, toolRegistry, forwarder, discoveredAt: Date.now() };
    this.userCache.set(sessionToken, entry);

    if (toolRegistry.size > 0) {
      const freshTools: CachedTool[] = Array.from(toolRegistry.values()).map((e) => ({
        name: e.originalName,
        description: e.toolDef.description!,
        inputSchema: e.toolDef.inputSchema as Record<string, unknown>,
        serverName: e.serverName,
      }));
      void saveToolsCache(freshTools);  // fire-and-forget, non-blocking
      this.rebuildSharedRegistry(freshTools);
      log(`Discovery complete: ${toolRegistry.size} WorkIQ + ${SHAREPOINT_TOOLS.length} SharePoint + 1 OCR tools`);
    }

    return entry;
  }

  private extractToken(req: express.Request): string | null {
    const urlToken = req.query.token as string | undefined;
    if (urlToken) return urlToken;
    const authHeader = req.headers.authorization;
    if (authHeader?.startsWith("Bearer ")) return authHeader.slice(7).trim();
    return null;
  }

  async start(): Promise<void> {
    const app = express();
    app.use(express.json());
    app.use(express.urlencoded({ extended: false }));

    app.use((_req, res, next) => {
      res.header("Access-Control-Allow-Origin", "*");
      res.header("Access-Control-Allow-Methods", "GET, POST, DELETE, HEAD, OPTIONS");
      res.header("Access-Control-Allow-Headers", "Content-Type, mcp-session-id, Authorization, Accept");
      if (_req.method === "OPTIONS") { res.sendStatus(200); return; }
      next();
    });

    const sessions: Record<string, SSEServerTransport> = {};

    app.get("/health", (_req, res) => {
      res.json({
        status: "ok",
        server: "agent365-bridge",
        cachedTools: this.sharedToolRegistry.size + SHAREPOINT_TOOLS.length + 2,
        activeUsers: this.userCache.size,
      });
    });

    const serverBaseUrl = process.env.SERVER_BASE_URL ||
      "https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io";

    if (this.config.tenantId && this.config.clientId && this.config.clientSecret) {
      registerOAuthEndpoints(app, {
        tenantId: this.config.tenantId,
        clientId: this.config.clientId,
        clientSecret: this.config.clientSecret,
        serverBaseUrl,
        mcpScope: this.config.mcpPlatformAuthScope,
      });
    } else {
      log("Warning: OAuth not configured");
    }

    app.head("/mcp", (_req, res) => {
      res.set("MCP-Protocol-Version", "2025-06-18");
      res.sendStatus(200);
    });

    app.post("/mcp", async (req, res) => {
      log(`MCP request: ${req.body?.method ?? "unknown"}`);

      const sessionToken = this.extractToken(req);
      if (!sessionToken) { await this.serveAnonymous(req, res); return; }

      const userToken = await getUserToken(sessionToken);
      const graphToken = await getGraphToken(sessionToken);
      const email = getTokenEmail(sessionToken);

      if (!userToken) {
        log("Session token expired or not found — sending WWW-Authenticate to trigger re-auth");
        res.status(401)
          .set("WWW-Authenticate", `Bearer realm="${serverBaseUrl}", error="invalid_token"`)
          .json({
            jsonrpc: "2.0",
            error: { code: -32001, message: "Session expired. Reconnecting..." },
            id: null,
          });
        return;
      }

      log(`Request from: ${email ?? sessionToken.substring(0, 8) + "..."}${graphToken ? " [Graph ✓]" : ""}`);

      const userCacheEntry = await this.getUserCache(sessionToken, userToken);
      const toolRegistry = userCacheEntry?.toolRegistry ?? this.sharedToolRegistry;
      const forwarder = userCacheEntry?.forwarder ?? null;

      try {
        const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
        const sessionServer = new Server(
          { name: "agent365-bridge", version: "1.0.0" },
          { capabilities: { tools: {} } }
        );

        sessionServer.setRequestHandler(ListToolsRequestSchema, async () => ({
          tools: [
            ...Array.from(toolRegistry.values()).map((e) => {
              // Patch descriptions of email composition tools to instruct the LLM
              // that it must call GetUserEmailSignatureStyle before composing emails.
              if (
                EMAIL_TOOLS_REQUIRING_SIGNATURE.has(e.uniqueName) ||
                EMAIL_TOOLS_REQUIRING_SIGNATURE.has(e.originalName)
              ) {
                return {
                  ...e.toolDef,
                  description:
                    (e.toolDef.description ?? "") +
                    "\n\nIMPORTANT: You MUST call GetUserEmailSignatureStyle BEFORE using this tool. Use the signature found there and append it verbatim to the bottom of the email body.",
                };
              }
              return e.toolDef;
            }),
            ...SHAREPOINT_TOOLS,
            OCR_TOOL,
            SIGNATURE_STYLE_TOOL,
          ],
        }));

        sessionServer.setRequestHandler(CallToolRequestSchema, async (request) => {
          const { name, arguments: args } = request.params;
          let typedArgs = (args ?? {}) as Record<string, unknown>;

          if (SHAREPOINT_TOOL_NAMES.has(name)) {
            const token = graphToken ?? userToken;
            const handler = new SharePointToolHandler(token);
            return handler.handleToolCall(name, typedArgs);
          }

          if (OCR_TOOL_NAMES.has(name)) {
            const token = graphToken ?? userToken;
            const handler = new OcrToolHandler(token);
            return handler.handleToolCall(name, typedArgs);
          }

          // ── Signature tool (dual-purpose: save if name given, read otherwise) ──
          if (name === SIGNATURE_STYLE_TOOL_NAME || name === SET_SIGNATURE_TOOL_NAME) {
            if (!graphToken) {
              return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate." }] };
            }
            const result = await handleGetSignatureStyle(graphToken, typedArgs);
            return { content: [{ type: "text", text: result }] };
          }

          const entry = toolRegistry.get(name);
          if (!entry) return { content: [{ type: "text", text: `Error: Tool '${name}' not found.` }] };
          if (!forwarder) return { content: [{ type: "text", text: "Tools still loading. Please try again in a moment." }] };

          const targetServer = userCacheEntry!.servers.find((s) => s.config.mcpServerName === entry.serverName);
          if (!targetServer) return { content: [{ type: "text", text: `Error: Server '${entry.serverName}' not found.` }] };

          // ── Auto-inject signature for outgoing email tools ──────────────────
          if (EMAIL_TOOLS_REQUIRING_SIGNATURE.has(name) && graphToken) {
            try {
              const signature = await getUserSignature(graphToken);
              if (signature) {
                const bodyKey = "body" in typedArgs ? "body"
                  : "emailBody" in typedArgs ? "emailBody"
                  : null;
                const bodyTypeKey = "bodyType" in typedArgs ? "bodyType"
                  : "contentType" in typedArgs ? "contentType"
                  : null;
                if (bodyKey && typeof typedArgs[bodyKey] === "string") {
                  const bodyType = bodyTypeKey ? (typedArgs[bodyTypeKey] as string ?? "text") : "text";
                  // appendSignature always returns HTML — update bodyType accordingly
                  typedArgs = { ...typedArgs, [bodyKey]: appendSignature(typedArgs[bodyKey] as string, bodyType, signature) };
                  if (bodyTypeKey) {
                    typedArgs = { ...typedArgs, [bodyTypeKey]: "html" };
                  }
                  log(`Signature auto-injected into ${name}`);
                }
              }
            } catch (sigErr) {
              log(`Signature auto-inject failed (non-fatal): ${sigErr}`);
            }
          }

          return await forwarder.callTool(entry.originalName, typedArgs, targetServer);
        });

        res.on("close", () => { transport.close(); sessionServer.close(); });
        await sessionServer.connect(transport);
        await transport.handleRequest(req, res, req.body);

      } catch (err) {
        log(`MCP request error: ${err}`);
        if (!res.headersSent) {
          res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal error" }, id: null });
        }
      }
    });

    app.get("/mcp", (_req, res) => {
      res.set("Allow", "POST");
      res.status(405).json({ error: "Method Not Allowed. Use POST." });
    });

    app.delete("/mcp", (_req, res) => {
      res.status(405).json({ error: "Method Not Allowed. Stateless mode." });
    });

    app.get("/sse", async (_req, res) => {
      log("New SSE connection");
      const transport = new SSEServerTransport("/messages", res);
      sessions[transport.sessionId] = transport;
      res.on("close", () => { delete sessions[transport.sessionId]; });
      const sseServer = new Server(
        { name: "agent365-bridge", version: "1.0.0" },
        { capabilities: { tools: {} } }
      );
      sseServer.setRequestHandler(ListToolsRequestSchema, async () => ({
        tools: [
          ...Array.from(this.sharedToolRegistry.values()).map((e) => e.toolDef),
          ...SHAREPOINT_TOOLS,
          OCR_TOOL,
          SIGNATURE_STYLE_TOOL,
          SET_SIGNATURE_TOOL,
        ],
      }));
      await sseServer.connect(transport);
    });

    app.post("/messages", async (req, res) => {
      const transport = sessions[req.query.sessionId as string];
      if (transport) await transport.handlePostMessage(req, res, req.body);
      else res.status(400).json({ error: "Session not found" });
    });

    setInterval(() => {
      const now = Date.now();
      for (const [token, entry] of this.userCache.entries()) {
        if (now - entry.discoveredAt > DISCOVERY_TTL_MS * 2) this.userCache.delete(token);
      }
    }, DISCOVERY_TTL_MS);

    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
      log(`MCP server listening on port ${PORT}`);
      log(`Login:  http://localhost:${PORT}/login`);
      log(`Health: http://localhost:${PORT}/health`);
    });
  }

  private async serveAnonymous(req: express.Request, res: express.Response): Promise<void> {
    try {
      const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
      const server = new Server({ name: "agent365-bridge", version: "1.0.0" }, { capabilities: { tools: {} } });
      server.setRequestHandler(ListToolsRequestSchema, async () => ({
        tools: [
          ...Array.from(this.sharedToolRegistry.values()).map((e) => e.toolDef),
          ...SHAREPOINT_TOOLS,
          OCR_TOOL,
          SIGNATURE_STYLE_TOOL,
        ],
      }));
      server.setRequestHandler(CallToolRequestSchema, async () => ({
        content: [{ type: "text", text: "Please visit /login to get your personal URL." }],
      }));
      res.on("close", () => { transport.close(); server.close(); });
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch {
      if (!res.headersSent) res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal error" }, id: null });
    }
  }
}

function sanitizeSchema(schema: Record<string, unknown>): Record<string, unknown> {
  const result = { ...schema };
  if (Array.isArray(result.allOf)) {
    const subSchemas = result.allOf as Record<string, unknown>[];
    for (const sub of subSchemas) {
      if (sub.properties) result.properties = { ...(result.properties as Record<string, unknown> ?? {}), ...(sub.properties as Record<string, unknown>) };
      if (Array.isArray(sub.required)) {
        const existing = Array.isArray(result.required) ? result.required as string[] : [];
        result.required = [...new Set([...existing, ...(sub.required as string[])])];
      }
    }
    delete result.allOf;
    if (!result.type) result.type = "object";
  }
  for (const keyword of ["oneOf", "anyOf"] as const) {
    if (Array.isArray(result[keyword])) {
      const variants = result[keyword] as Record<string, unknown>[];
      if (variants.length > 0) {
        const first = variants[0];
        if (first.properties) result.properties = { ...(result.properties as Record<string, unknown> ?? {}), ...(first.properties as Record<string, unknown>) };
        if (Array.isArray(first.required)) {
          const existing = Array.isArray(result.required) ? result.required as string[] : [];
          result.required = [...new Set([...existing, ...(first.required as string[])])];
        }
      }
      delete result[keyword];
      if (!result.type) result.type = "object";
    }
  }
  if (result.properties && typeof result.properties === "object") {
    const props = result.properties as Record<string, Record<string, unknown>>;
    const cleaned: Record<string, unknown> = {};
    for (const [key, val] of Object.entries(props)) {
      cleaned[key] = typeof val === "object" && val !== null ? sanitizeSchema(val) : val;
    }
    result.properties = cleaned;
  }
  return result;
}