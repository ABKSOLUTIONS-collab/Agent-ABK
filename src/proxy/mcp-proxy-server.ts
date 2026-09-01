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
import { getUserToken, getGraphToken, getStoredEmail, pruneExpiredTokens } from "../auth/user-token-store";
import { registerOAuthEndpoints } from "../auth/oauth-handler";
import { SHAREPOINT_TOOLS, SHAREPOINT_TOOL_NAMES, SharePointToolHandler, fetchDriveItem, createDriveItemLink, FetchedDriveItem } from "../tools/sharepoint-tools";
import { OCR_TOOL, OCR_TOOL_NAMES, OcrToolHandler } from "../tools/ocr-tool";
import { EMAIL_GRAPH_TOOLS, EMAIL_GRAPH_TOOL_NAMES, EmailGraphToolHandler } from "../tools/email-graph-tools";
import { CALENDAR_GRAPH_TOOLS, CALENDAR_GRAPH_TOOL_NAMES, CalendarGraphToolHandler } from "../tools/calendar-graph-tools";
import { WORD_GRAPH_TOOLS, WORD_GRAPH_TOOL_NAMES, WordGraphToolHandler } from "../tools/word-graph-tools";
import { EXCEL_GRAPH_TOOLS, EXCEL_GRAPH_TOOL_NAMES, ExcelGraphToolHandler } from "../tools/excel-graph-tools";
import { TEAMS_GRAPH_TOOLS, TEAMS_GRAPH_TOOL_NAMES, TeamsGraphToolHandler } from "../tools/teams-graph-tools";
import { KNOWLEDGE_GRAPH_TOOLS, KNOWLEDGE_GRAPH_TOOL_NAMES, KnowledgeGraphToolHandler } from "../tools/knowledge-graph-tools";
import { POWERPOINT_TOOLS, POWERPOINT_TOOL_NAMES, PowerPointToolHandler } from "../tools/powerpoint-tools";
import {
  SIGNATURE_STYLE_TOOL,
  SIGNATURE_STYLE_TOOL_NAME,
  SET_SIGNATURE_TOOL,
  SET_SIGNATURE_TOOL_NAME,
  EMAIL_TOOLS_REQUIRING_SIGNATURE,
  handleGetSignatureStyle,
  handleSetSignature,
} from "../tools/signature-style-tool";
import { getUserSignatureStrict, appendSignature, LOGO_BASE64 } from "../tools/signature-service";
import { ServerDiscovery } from "../discovery/server-discovery";
import { AppConfig } from "../config/types";
import express from "express";
import cookieParser from "cookie-parser";
import { registerOrgAdminEndpoints } from "../admin/org-admin";

function log(message: string): void {
  process.stderr.write(`[agent365-bridge] ${message}\n`);
}

// Returned instead of sending/drafting when the user has no saved signature
// yet, with isError:true so it's surfaced as a genuine tool failure rather
// than free-form text the calling model has to interpret. Deliberately
// phrased as a plain fact about current state, with zero imperative verbs
// ("call X", "ask the user", "must") — any such directive inside a tool
// result reads exactly like an injected instruction, and two earlier, more
// forceful wordings of this were both (correctly) flagged as suspicious and
// refused. The model already sees GetUserEmailSignatureStyle in its own
// tool list with a description explaining what it's for; it doesn't need to
// be told in-band to use it.
const SIGNATURE_REQUIRED_MESSAGE =
  "Precondition not met: this Microsoft 365 account has no email signature configured " +
  "(Agent365-Bridge/signature.txt not found in its OneDrive). Email not sent or drafted.";

// Tool names that already have a hardcoded Graph-API handler and are always
// listed via their static tool arrays — used to keep discovered (dynamic)
// tools from being listed a second time under the same name.
const HARDCODED_TOOL_NAMES = new Set<string>([
  ...SHAREPOINT_TOOL_NAMES,
  ...OCR_TOOL_NAMES,
  ...EMAIL_GRAPH_TOOL_NAMES,
  ...CALENDAR_GRAPH_TOOL_NAMES,
  ...WORD_GRAPH_TOOL_NAMES,
  ...EXCEL_GRAPH_TOOL_NAMES,
  ...TEAMS_GRAPH_TOOL_NAMES,
  ...KNOWLEDGE_GRAPH_TOOL_NAMES,
  ...POWERPOINT_TOOL_NAMES,
  SIGNATURE_STYLE_TOOL_NAME,
  SET_SIGNATURE_TOOL_NAME,
]);

// Discovered (not hardcoded) tools that resolve a SharePoint site by name —
// these come from the live Agent 365 gateway and do literal/partial matching
// with no fuzzy fallback. See the site-name search fallback below.
const SITE_NAME_SEARCH_TOOL_NAMES = new Set<string>(["searchSitesByName"]);

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

// ── Discoverable tool surface ────────────────────────────────────────────────
// Rather than listing all ~70 tools (hardcoded + discovered) up front — which
// means re-sending every schema as input tokens on every single message,
// whether or not that turn needs SharePoint/Excel/Teams/etc. — only these two
// meta-tools are advertised. find_tools searches the full catalog by keyword
// and returns matching schemas; call_tool then invokes one by exact name.
// Nothing is removed: every real tool is still reachable, just addressed
// indirectly instead of always being in context.
const FIND_TOOLS_TOOL: Tool = {
  name: "find_tools",
  description:
    "Search the full Microsoft 365 tool catalog (mail, calendar, Word, Excel, Teams, " +
    "SharePoint, OneDrive, knowledge search, OCR, email signature, and more) by keyword. " +
    "Only find_tools and call_tool are listed up front to save context — this is how you " +
    "discover everything else. Returns matching tools with their full parameter schema. " +
    "Call this whenever you need a capability that isn't already visible to you, then " +
    "invoke the result via call_tool.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description:
          "Keywords describing what you want to do, e.g. 'reply to an email', " +
          "'list SharePoint folder', 'create calendar event', 'search files'.",
      },
    },
    required: ["query"],
  },
};

const CALL_TOOL_TOOL: Tool = {
  name: "call_tool",
  description:
    "Invoke any tool found via find_tools. Pass the exact tool name and arguments " +
    "matching the schema find_tools returned for it.",
  inputSchema: {
    type: "object",
    properties: {
      name: { type: "string", description: "Exact tool name, as returned by find_tools." },
      arguments: { type: "object", description: "Arguments matching that tool's inputSchema." },
    },
    required: ["name"],
  },
};

// Loose shape for the searchable catalog — OCR_TOOL/SIGNATURE_STYLE_TOOL
// aren't explicitly typed as Tool in their source files, and these are only
// ever rendered into a find_tools text listing, never handed back to the SDK,
// so a strict Tool[] return type here would reject them for no benefit.
interface ToolLike {
  name: string;
  description?: string;
  inputSchema: unknown;
}

function allToolDefs(toolRegistry: Map<string, ToolRegistryEntry>): ToolLike[] {
  return [
    ...Array.from(toolRegistry.values()).map((e) => e.toolDef),
    ...SHAREPOINT_TOOLS,
    ...EMAIL_GRAPH_TOOLS,
    ...CALENDAR_GRAPH_TOOLS,
    ...WORD_GRAPH_TOOLS,
    ...EXCEL_GRAPH_TOOLS,
    ...TEAMS_GRAPH_TOOLS,
    ...KNOWLEDGE_GRAPH_TOOLS,
    ...POWERPOINT_TOOLS,
    OCR_TOOL,
    SIGNATURE_STYLE_TOOL,
  ];
}

function searchTools(toolRegistry: Map<string, ToolRegistryEntry>, query: string, limit = 12): ToolLike[] {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  const all = allToolDefs(toolRegistry);
  if (terms.length === 0) return all.slice(0, limit);
  const scored = all
    .map((t) => {
      const haystack = `${t.name} ${t.description ?? ""}`.toLowerCase();
      const score = terms.reduce((s, term) => s + (haystack.includes(term) ? 1 : 0), 0);
      return { t, score };
    })
    .filter((x) => x.score > 0);
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, limit).map((x) => x.t);
}

interface ToolCallContext {
  graphToken: string | null;
  email: string | undefined;
  toolRegistry: Map<string, ToolRegistryEntry>;
  forwarder: ToolForwarder | null;
  userCacheEntry: UserDiscoveryCache | null;
}

// Graph rejects a plain fileAttachment POST above ~3 MB, so anything larger has
// to go through a resumable upload session. 3 MB is the documented boundary;
// staying just under it keeps the simple path for the common case.
const DIRECT_ATTACH_LIMIT = 3 * 1024 * 1024;
const UPLOAD_CHUNK = 4 * 1024 * 1024;

async function attachFileToMessage(
  graphToken: string,
  msgId: string,
  file: FetchedDriveItem,
  // Which mailbox the draft lives in — the signed-in user's own, or a shared
  // one being sent as. Attachments must be posted to the same mailbox as the
  // draft, or Graph cannot find the message.
  mailboxBase = "https://graph.microsoft.com/v1.0/me"
): Promise<void> {
  const auth = { Authorization: `Bearer ${graphToken}` };

  if (file.buffer.length <= DIRECT_ATTACH_LIMIT) {
    const res = await fetch(`${mailboxBase}/messages/${msgId}/attachments`, {
      method: "POST",
      headers: { ...auth, "Content-Type": "application/json" },
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: file.name,
        contentType: file.contentType,
        contentBytes: file.buffer.toString("base64"),
      }),
    });
    if (!res.ok) throw new Error(`attach failed (${res.status}): ${await res.text()}`);
    return;
  }

  const sessRes = await fetch(`${mailboxBase}/messages/${msgId}/attachments/createUploadSession`, {
    method: "POST",
    headers: { ...auth, "Content-Type": "application/json" },
    body: JSON.stringify({
      AttachmentItem: { attachmentType: "file", name: file.name, size: file.buffer.length },
    }),
  });
  if (!sessRes.ok) throw new Error(`upload session failed (${sessRes.status}): ${await sessRes.text()}`);
  const { uploadUrl } = await sessRes.json() as { uploadUrl: string };

  // The upload URL is pre-authorised, so these PUTs deliberately carry no
  // Authorization header — Graph rejects the request if one is sent.
  const total = file.buffer.length;
  for (let start = 0; start < total; start += UPLOAD_CHUNK) {
    const end = Math.min(start + UPLOAD_CHUNK, total) - 1;
    const chunk = file.buffer.subarray(start, end + 1);
    const put = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${end}/${total}`,
      },
      body: new Uint8Array(chunk),
    });
    if (!put.ok && put.status !== 201 && put.status !== 202) {
      throw new Error(`chunk ${start}-${end} failed (${put.status}): ${await put.text()}`);
    }
  }
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
      log(`Loaded ${this.sharedToolRegistry.size} tools from Blob Storage cache`);
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
        // These names already have a hardcoded Graph-API handler and are
        // always listed via their static tool arrays — skip them here so
        // discovery doesn't list the same tool name twice.
        if (HARDCODED_TOOL_NAMES.has(t.name)) continue;
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

    // ── Serve static assets (logo etc.) ──────────────────────────────────────
    app.use("/assets", express.static(
      require("path").join(__dirname, "../../assets"),
      { maxAge: "7d", immutable: true }
    ));

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

    app.use(cookieParser());
    registerOrgAdminEndpoints(app);

    app.head("/mcp", (_req, res) => {
      res.set("MCP-Protocol-Version", "2025-06-18");
      res.sendStatus(200);
    });

    app.post("/mcp", async (req, res) => {
      log(`MCP request: ${req.body?.method ?? "unknown"}`);

      const sessionToken = this.extractToken(req);
      if (!sessionToken) {
        log("No session token provided — sending WWW-Authenticate to trigger OAuth");
        res.status(401)
          .set("WWW-Authenticate", `Bearer realm="${serverBaseUrl}", error="invalid_token"`)
          .json({
            jsonrpc: "2.0",
            error: { code: -32001, message: "Authentication required. Please connect via OAuth." },
            id: null,
          });
        return;
      }

      const [userToken, graphToken, email] = await Promise.all([
        getUserToken(sessionToken),
        getGraphToken(sessionToken),
        getStoredEmail(sessionToken),
      ]);

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
          tools: [FIND_TOOLS_TOOL, CALL_TOOL_TOOL],
        }));

        sessionServer.setRequestHandler(CallToolRequestSchema, async (request) => {
          const { name, arguments: args } = request.params;
          const typedArgs = (args ?? {}) as Record<string, unknown>;
          const ctx: ToolCallContext = { graphToken, email, toolRegistry, forwarder, userCacheEntry };

          if (name === FIND_TOOLS_TOOL.name) {
            const query = String(typedArgs.query ?? "");
            const matches = searchTools(toolRegistry, query);
            if (matches.length === 0) {
              return { content: [{ type: "text", text: `No tools matched "${query}". Try broader or different keywords.` }] };
            }
            const listing = matches
              .map((t) => `### ${t.name}\n${t.description ?? ""}\nSchema: ${JSON.stringify(t.inputSchema)}`)
              .join("\n\n");
            return {
              content: [{
                type: "text",
                text: `Found ${matches.length} matching tool(s). Call each via call_tool with the exact name and arguments matching its schema below:\n\n${listing}`,
              }],
            };
          }

          if (name === CALL_TOOL_TOOL.name) {
            const innerName = String(typedArgs.name ?? "");
            const innerArgs = (typedArgs.arguments ?? {}) as Record<string, unknown>;
            if (!innerName) {
              return { content: [{ type: "text", text: "Error: call_tool requires a 'name' field — use find_tools first to get the exact tool name and schema." }], isError: true };
            }
            return executeToolCall(innerName, innerArgs, ctx);
          }

          // Backward-compat: a caller that already knows a real tool name
          // (e.g. remembered from earlier in the conversation) can still
          // invoke it directly without going through call_tool.
          return executeToolCall(name, typedArgs, ctx);
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
          ...EMAIL_GRAPH_TOOLS,
          ...CALENDAR_GRAPH_TOOLS,
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
      log(`Login:          http://localhost:${PORT}/login`);
      log(`Admin consent:  http://localhost:${PORT}/admin-consent  (run ONCE by an Azure AD admin)`);
      log(`Health:         http://localhost:${PORT}/health`);
    });
  }

}

// The actual per-tool dispatch logic, reached either directly (backward-compat
// path for a caller that already knows a real tool name) or via call_tool.
async function executeToolCall(
  name: string,
  argsIn: Record<string, unknown>,
  ctx: ToolCallContext
): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
  const { graphToken, email, toolRegistry, forwarder, userCacheEntry } = ctx;
  let typedArgs = argsIn;

  if (SHAREPOINT_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new SharePointToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  if (OCR_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new OcrToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── CreateDraftMessage: inject signature + CID logo before creating ──
  if (name === "CreateDraftMessage" && graphToken) {
    let signature: string | null = null;
    try {
      signature = await getUserSignatureStrict(graphToken);
    } catch (e) { log(`Signature check failed: ${e}`); }

    if (!signature) {
      return { content: [{ type: "text", text: SIGNATURE_REQUIRED_MESSAGE }], isError: true };
    }

    const rawBody  = String(typedArgs.body ?? typedArgs.emailBody ?? "");
    const bodyType = String(typedArgs.contentType ?? typedArgs.bodyType ?? "text");
    typedArgs = { ...typedArgs, body: appendSignature(rawBody, bodyType, signature), contentType: "HTML" };
    log("Signature injected into CreateDraftMessage");

    const handler = new EmailGraphToolHandler(graphToken);
    const result   = await handler.handleToolCall(name, typedArgs);

    // Add CID logo attachment to the newly created draft
    if (LOGO_BASE64) {
      try {
        const text  = (result.content?.[0] as { text?: string })?.text ?? "";
        const match = text.match(/ID:\s*([A-Za-z0-9_\-=]+)/);
        if (match?.[1]) {
          await fetch(`https://graph.microsoft.com/v1.0/me/messages/${match[1]}/attachments`, {
            method: "POST",
            headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
            body: JSON.stringify({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: "abk-logo.png", contentType: "image/png",
              contentBytes: LOGO_BASE64, contentId: "abk-logo@abk", isInline: true,
            }),
          });
          log(`CID logo attachment added to draft ${match[1]}`);
        }
      } catch (e) { log(`CID attachment failed (non-fatal): ${e}`); }
    }
    return result;
  }

  // ── SendEmailWithAttachments: create draft + CID attachment + send ──
  // We can't call /me/sendMail directly (needs Mail.Send on graphToken).
  // Instead: create draft via Graph (Mail.ReadWrite ✓) → add CID attachment
  // → send via upstream SendDraftMessage tool (has Mail.Send ✓).
  // NOTE: must run BEFORE the generic EMAIL_GRAPH_TOOL_NAMES check below,
  // which would otherwise intercept this name first and skip signature/CID injection.
  if (name === "SendEmailWithAttachments" && graphToken && LOGO_BASE64) {
    const signature = await getUserSignatureStrict(graphToken).catch(() => null);
    if (!signature) {
      return { content: [{ type: "text", text: SIGNATURE_REQUIRED_MESSAGE }], isError: true };
    }
    try {
      const rawBody  = String(typedArgs.body ?? typedArgs.emailBody ?? "");
      const bodyType = String(typedArgs.contentType ?? typedArgs.bodyType ?? "text");

      // Sharing links are resolved into the body BEFORE the signature is
      // appended, so they read as part of the message rather than trailing
      // after the sign-off.
      let bodyWithLinks = rawBody;
      const linkRefs = Array.isArray(typedArgs.links) ? typedArgs.links : [];
      const linkedNames: string[] = [];
      if (linkRefs.length) {
        const rendered: string[] = [];
        for (const raw of linkRefs) {
          const spec = (typeof raw === "string" ? { ref: raw } : raw) as { ref?: string; siteId?: string; type?: string };
          if (!spec?.ref) continue;
          try {
            const link = await createDriveItemLink(
              graphToken, spec.ref, spec.siteId,
              spec.type === "edit" ? "edit" : "view"
            );
            rendered.push(`<li><a href="${link.url}">${link.name}</a></li>`);
            linkedNames.push(link.name);
          } catch (linkErr) {
            const m = linkErr instanceof Error ? linkErr.message : String(linkErr);
            log(`Link failed for '${spec.ref}': ${m}`);
            return {
              content: [{ type: "text", text: `Could not create a sharing link for '${spec.ref}': ${m}\nThe email was NOT sent.` }],
              isError: true,
            };
          }
        }
        if (rendered.length) {
          // Plain-text bodies are wrapped first; the message is sent as HTML
          // either way, so a bare newline would otherwise collapse.
          const bodyHtml = bodyType.toLowerCase() === "html"
            ? bodyWithLinks
            : bodyWithLinks.replace(/\n/g, "<br>");
          bodyWithLinks = `${bodyHtml}<ul>${rendered.join("")}</ul>`;
        }
      }

      const htmlBody = appendSignature(bodyWithLinks, linkRefs.length ? "html" : bodyType, signature);

      // Schema allows "to"/"cc"/"bcc" as either a single string or an
      // array of strings — normalize both before mapping to Graph shape.
      const buildRecipients = (val: unknown) =>
        (Array.isArray(val) ? val : val ? [val] : [])
          .map((r: unknown) => ({ emailAddress: { address: String(r) } }));

      // Sending as a shared mailbox (e.g. info@) is a normal, admin-granted
      // arrangement, not spoofing: Exchange independently enforces whether
      // this user actually holds Send As rights on that mailbox.
      const fromAddress = String(typedArgs.from ?? "").trim();
      const sendAsShared = !!fromAddress && fromAddress.toLowerCase() !== (email ?? "").toLowerCase();

      // The shared-mailbox path deliberately does NOT create a draft first.
      // Graph splits these permissions:
      //   POST /users/{x}/messages  (create a draft there) -> Mail.ReadWrite.Shared
      //   POST /users/{x}/sendMail  (send as it)           -> Mail.Send.Shared
      // Our tokens carry Mail.Send.Shared but not Mail.ReadWrite.Shared, so
      // building a draft in the shared mailbox returns 403 even for a user who
      // genuinely has Send As. sendMail takes the entire message in one
      // request instead, which is why attachments are inlined below.
      if (sendAsShared) {
        const inline: Record<string, unknown>[] = [];
        if (LOGO_BASE64) {
          inline.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: "abk-logo.png", contentType: "image/png",
            contentBytes: LOGO_BASE64, contentId: "abk-logo@abk", isInline: true,
          });
        }

        const attachRefsShared = Array.isArray(typedArgs.attachments) ? typedArgs.attachments : [];
        const attachedShared: string[] = [];
        let inlineBytes = 0;
        for (const raw of attachRefsShared) {
          const spec = (typeof raw === "string" ? { ref: raw } : raw) as { ref?: string; siteId?: string };
          if (!spec?.ref) continue;
          try {
            const file = await fetchDriveItem(graphToken, spec.ref, spec.siteId);
            inlineBytes += file.buffer.length;
            // sendMail carries everything in a single request, which Graph caps
            // at roughly 4 MB. The resumable upload used for normal sends is
            // not available here because there is no draft to upload into.
            if (inlineBytes > DIRECT_ATTACH_LIMIT) {
              return {
                content: [{ type: "text", text: `'${file.name}' makes the attachments too large to send from a shared mailbox (limit is about 3 MB in total). Send it from your own mailbox instead, or share a link to the file.` }],
                isError: true,
              };
            }
            inline.push({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: file.name,
              contentType: file.contentType,
              contentBytes: file.buffer.toString("base64"),
            });
            attachedShared.push(file.name);
          } catch (attErr) {
            const m = attErr instanceof Error ? attErr.message : String(attErr);
            return { content: [{ type: "text", text: `Could not attach '${spec.ref}': ${m}. The email was NOT sent.` }], isError: true };
          }
        }

        const sharedRes = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromAddress)}/sendMail`, {
          method: "POST",
          headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
          body: JSON.stringify({
            message: {
              subject:       String(typedArgs.subject ?? ""),
              body:          { contentType: "HTML", content: htmlBody },
              toRecipients:  buildRecipients(typedArgs.to),
              ccRecipients:  buildRecipients(typedArgs.cc),
              bccRecipients: buildRecipients(typedArgs.bcc),
              ...(inline.length ? { attachments: inline } : {}),
            },
            // Keeps a copy in the shared mailbox's Sent Items, where the
            // colleagues who share it expect to find what was sent.
            saveToSentItems: true,
          }),
        });

        if (sharedRes.ok || sharedRes.status === 202) {
          log(`Email sent as ${fromAddress}${attachedShared.length ? ` with ${attachedShared.length} attachment(s)` : ""}${linkedNames.length ? ` and ${linkedNames.length} link(s)` : ""}`);
          const suffix = attachedShared.length ? ` Attached: ${attachedShared.join(", ")}.` : "";
          const links = linkedNames.length ? ` Linked: ${linkedNames.join(", ")}.` : "";
          return { content: [{ type: "text", text: `Email sent successfully from '${fromAddress}'.${suffix}${links}` }] };
        }

        const detail = await sharedRes.text().catch(() => "");
        const denied = sharedRes.status === 403 || sharedRes.status === 404;
        return {
          content: [{
            type: "text",
            text: denied
              ? `Could not send as '${fromAddress}': Exchange refused it, which means the signed-in account does not hold Send As rights on that mailbox, or the mailbox does not exist. An Exchange administrator can grant Send As on '${fromAddress}'. (Graph ${sharedRes.status})`
              : `Could not send as '${fromAddress}' (${sharedRes.status}): ${detail}`,
          }],
          isError: true,
        };
      }

      const mailboxBase = "https://graph.microsoft.com/v1.0/me";

      // 1️⃣ Create draft
      const draftResp = await fetch(`${mailboxBase}/messages`, {
        method: "POST",
        headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          subject:       String(typedArgs.subject ?? ""),
          body:          { contentType: "HTML", content: htmlBody },
          toRecipients:  buildRecipients(typedArgs.to),
          ccRecipients:  buildRecipients(typedArgs.cc),
          bccRecipients: buildRecipients(typedArgs.bcc),
        }),
      });

      if (!draftResp.ok) {
        log(`Create draft failed: ${draftResp.status} — falling back to upstream`);
      } else {
        const draft = await draftResp.json() as { id: string };
        const msgId = draft.id;

        // 2️⃣ Add inline CID attachment
        await fetch(`${mailboxBase}/messages/${msgId}/attachments`, {
          method: "POST",
          headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
          body: JSON.stringify({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name:         "abk-logo.png",
            contentType:  "image/png",
            contentBytes: LOGO_BASE64,
            contentId:    "abk-logo@abk",
            isInline:     true,
          }),
        });

        // 2️⃣b Attach any requested OneDrive / SharePoint files.
        // Done here, on the draft, rather than through /me/sendMail: sendMail
        // takes the whole message in one request, so a large file would have to
        // be inlined into that payload, while a draft accepts attachments one
        // at a time and supports resumable upload for anything over 3 MB.
        const attachRefs = Array.isArray(typedArgs.attachments) ? typedArgs.attachments : [];
        const attachedNames: string[] = [];
        for (const raw of attachRefs) {
          const spec = (typeof raw === "string" ? { ref: raw } : raw) as { ref?: string; siteId?: string };
          if (!spec?.ref) continue;
          try {
            const file = await fetchDriveItem(graphToken, spec.ref, spec.siteId);
            await attachFileToMessage(graphToken, msgId, file, mailboxBase);
            attachedNames.push(file.name);
          } catch (attErr) {
            // Abort rather than send a mail that silently lacks the very file
            // it was about. The draft is left in place so nothing is lost.
            const m = attErr instanceof Error ? attErr.message : String(attErr);
            log(`Attachment failed for '${spec.ref}': ${m}`);
            return {
              content: [{ type: "text", text: `Could not attach '${spec.ref}': ${m}\nThe email was NOT sent; a draft remains in the mailbox (id: ${msgId}).` }],
              isError: true,
            };
          }
        }

        // 3️⃣ Send directly via Graph API (no Copilot license needed)
        const sendRes = await fetch(`${mailboxBase}/messages/${msgId}/send`, {
          method: "POST",
          headers: { Authorization: `Bearer ${graphToken}` },
        });
        if (sendRes.ok || sendRes.status === 202) {
          // Only the user's own mailbox reaches here; the shared-mailbox path
          // returns above.
          log(`Email sent via Graph API direct send (msg: ${msgId})${attachedNames.length ? ` with ${attachedNames.length} attachment(s)` : ""}${linkedNames.length ? ` and ${linkedNames.length} link(s)` : ""}`);
          const suffix = attachedNames.length ? ` Attached: ${attachedNames.join(", ")}.` : "";
          const links = linkedNames.length ? ` Linked: ${linkedNames.join(", ")}.` : "";
          return { content: [{ type: "text", text: `Email sent successfully.${suffix}${links}` }] };
        }
        const sendErr = await sendRes.text().catch(() => "");
        log(`Graph direct send failed (${sendRes.status}): ${sendErr} — trying upstream`);

        // Fallback: upstream SendDraftMessage. Look this up directly on the
        // discovered servers (not the deduped toolRegistry, which excludes
        // this name since it also has a hardcoded handler) so forwarding
        // still works when the direct Graph send above fails for some
        // reason other than the bug this was fixed for.
        if (forwarder && userCacheEntry) {
          for (const s of userCacheEntry.servers) {
            const dynTool = s.tools.find(t => t.name === "SendDraftMessage");
            if (dynTool) {
              const result = await forwarder.callTool("SendDraftMessage", { messageId: msgId }, s);
              log(`Email sent via draft+CID+upstream-send (msg: ${msgId})`);
              return result;
            }
          }
        }

        log(`Email draft created with CID logo (msg: ${msgId}) — could not auto-send`);
        return { content: [{ type: "text", text: `Draft created with your signature, but could not send automatically (${sendErr || sendRes.status}). Please send it from Outlook (id: ${msgId})` }] };
      }
    } catch (sigErr) {
      log(`Draft+CID send failed (non-fatal): ${sigErr} — falling back to upstream`);
    }
  }

  // ── SendDraftMessage: add CID attachment then send directly via Graph ──
  // NOTE: must also run BEFORE the generic EMAIL_GRAPH_TOOL_NAMES check below.
  if (name === "SendDraftMessage" && graphToken) {
    try {
      const msgId = String(typedArgs.messageId ?? typedArgs.draftId ?? typedArgs.id ?? "");
      if (msgId) {
        if (LOGO_BASE64) {
          await fetch(`https://graph.microsoft.com/v1.0/me/messages/${msgId}/attachments`, {
            method: "POST",
            headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
            body: JSON.stringify({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: "abk-logo.png", contentType: "image/png",
              contentBytes: LOGO_BASE64, contentId: "abk-logo@abk", isInline: true,
            }),
          });
          log(`CID attachment added to draft ${msgId} before send`);
        }
        const sendRes = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${msgId}/send`, {
          method: "POST",
          headers: { Authorization: `Bearer ${graphToken}` },
        });
        if (sendRes.ok || sendRes.status === 202) {
          log(`Draft sent directly via Graph API (msg: ${msgId})`);
          return { content: [{ type: "text", text: "Email sent successfully." }] };
        }
        const sendErr = await sendRes.text().catch(() => "");
        log(`Graph direct send failed (${sendRes.status}): ${sendErr} — falling back to upstream`);
      }
    } catch (e) {
      log(`Graph send attempt failed (non-fatal): ${e} — falling back to upstream`);
    }
  }

  // ── Email tools: Graph API directly (no Copilot license required) ──
  if (EMAIL_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new EmailGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Calendar tools: Graph API directly (no Copilot license required) ──
  if (CALENDAR_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new CalendarGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Word document tools: Graph API + DOCX (no Copilot license required) ──
  if (WORD_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new WordGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Excel workbook tools: Graph Excel API (no Copilot license required) ──
  if (EXCEL_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new ExcelGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Teams meeting tools: Graph API (no Copilot license required) ──
  if (TEAMS_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new TeamsGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Knowledge / Search tools: Graph Search API ──────────────────
  if (KNOWLEDGE_GRAPH_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new KnowledgeGraphToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── PowerPoint: read deck text straight from the .pptx package ──
  if (POWERPOINT_TOOL_NAMES.has(name)) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate at /login." }], isError: true };
    }
    const handler = new PowerPointToolHandler(graphToken);
    return handler.handleToolCall(name, typedArgs);
  }

  // ── Signature tool (dual-purpose: save if name given, read otherwise) ──
  if (name === SIGNATURE_STYLE_TOOL_NAME || name === SET_SIGNATURE_TOOL_NAME) {
    if (!graphToken) {
      return { content: [{ type: "text", text: "Graph token not available. Please re-authenticate." }], isError: true };
    }
    // Guard against saving someone else's details (e.g. a person the
    // caller has emailed before) as THIS account's signature — the
    // provided email, if any, must belong to the signed-in user.
    const providedEmail = typeof typedArgs.email === "string" ? typedArgs.email.trim().toLowerCase() : "";
    if (providedEmail && email && providedEmail !== email.toLowerCase()) {
      return { content: [{ type: "text", text: `This account is signed in as ${email}, but the email provided (${providedEmail}) belongs to someone else. A signature must be the signed-in user's own details, not another person's (e.g. an email recipient). Ask the user to confirm their own name/title/contact info and call this tool again with those — or omit the email field.` }], isError: true };
    }
    const result = await handleGetSignatureStyle(graphToken, typedArgs);
    return { content: [{ type: "text", text: result }] };
  }

  const entry = toolRegistry.get(name);
  if (!entry) return { content: [{ type: "text", text: `Error: Tool '${name}' not found. Use find_tools to look up the exact name.` }], isError: true };
  if (!forwarder) return { content: [{ type: "text", text: "Tools still loading. Please try again in a moment." }], isError: true };

  const targetServer = userCacheEntry!.servers.find((s) => s.config.mcpServerName === entry.serverName);
  if (!targetServer) return { content: [{ type: "text", text: `Error: Server '${entry.serverName}' not found.` }], isError: true };

  // ── Other email tools: require + inject signature into body ─────────
  if (EMAIL_TOOLS_REQUIRING_SIGNATURE.has(name) && graphToken && name !== "SendEmailWithAttachments") {
    const signature = await getUserSignatureStrict(graphToken).catch(() => null);
    if (!signature) {
      return { content: [{ type: "text", text: SIGNATURE_REQUIRED_MESSAGE }], isError: true };
    }
    const bodyKey = "body" in typedArgs ? "body" : "emailBody" in typedArgs ? "emailBody" : null;
    const bodyTypeKey = "bodyType" in typedArgs ? "bodyType" : "contentType" in typedArgs ? "contentType" : null;
    if (bodyKey && typeof typedArgs[bodyKey] === "string") {
      const bodyType = bodyTypeKey ? (typedArgs[bodyTypeKey] as string ?? "text") : "text";
      typedArgs = { ...typedArgs, [bodyKey]: appendSignature(typedArgs[bodyKey] as string, bodyType, signature), contentType: "HTML" };
      log(`Signature auto-injected into ${name}`);
    }
  }

  // ── Site-name search fallback: Microsoft's site-name search is a
  // literal/partial match with no spell-correction, so one typo in a
  // site name ("Exeperia" vs "Experia") returns zero results and dead-
  // ends the conversation. Microsoft Graph Search (query_federated_
  // knowledge) does relevance-ranked, typo-tolerant matching, so when
  // the exact-name lookup comes back empty, retry through it before
  // telling the user the site doesn't exist / they lack access.
  if (SITE_NAME_SEARCH_TOOL_NAMES.has(entry.originalName)) {
    const primary = await forwarder.callTool(entry.originalName, typedArgs, targetServer);
    const primaryText = (primary.content?.[0] as { text?: string })?.text ?? "";
    const cameBackEmpty = /"value"\s*:\s*\[\s*\]/.test(primaryText) || !primaryText.trim();
    if (cameBackEmpty && graphToken) {
      const query = String(typedArgs.search ?? typedArgs.query ?? typedArgs.name ?? "");
      log(`${entry.originalName} returned no results for "${query}" — falling back to query_federated_knowledge`);
      const fallback = await new KnowledgeGraphToolHandler(graphToken).handleToolCall(
        "query_federated_knowledge",
        { query, entityTypes: ["site"] }
      );
      const fallbackText = (fallback.content?.[0] as { text?: string })?.text ?? "";
      return {
        content: [{
          type: "text",
          text: `No exact site name match for "${query}". Falling back to Microsoft Search, which tolerates typos and partial names:\n\n${fallbackText}`,
        }],
      };
    }
    return primary;
  }

  return await forwarder.callTool(entry.originalName, typedArgs, targetServer);
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