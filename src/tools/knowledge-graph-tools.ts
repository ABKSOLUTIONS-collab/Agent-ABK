/**
 * knowledge-graph-tools.ts
 *
 * Federated Knowledge tools.
 *
 * - query_federated_knowledge  → Microsoft Graph Search API (searches mail, files,
 *   SharePoint, Teams, and people across the user's M365 tenant — no Copilot needed)
 *
 * - configure / ingest / delete / retrieve_federated_knowledge → WorkIQ platform
 *   admin features that configure external data connectors inside Microsoft Copilot.
 *   These require a Copilot license AND admin rights. Regular employees never need them.
 *   Stubs below explain what they do and return a clear message.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [knowledge-graph] ${msg}\n`);
}

interface GraphFetchResult {
  ok: boolean;
  status: number;
  body: unknown;
}

async function gf(token: string, path: string, options: RequestInit = {}): Promise<GraphFetchResult> {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers ?? {}),
    },
  });
  const ct = res.headers.get("content-type") ?? "";
  const body = ct.includes("json") || ct.includes("odata")
    ? await res.json().catch(() => null)
    : await res.text().catch(() => "");
  return { ok: res.ok, status: res.status, body };
}

// ── Tool names ────────────────────────────────────────────────────────────────

export const KNOWLEDGE_GRAPH_TOOL_NAMES = new Set([
  "query_federated_knowledge",
  "configure_federated_knowledge",
  "delete_federated_knowledge",
  "ingest_federated_knowledge",
  "retrieve_federated_knowledge",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const KNOWLEDGE_GRAPH_TOOLS: Tool[] = [
  {
    name: "query_federated_knowledge",
    description:
      "Search content across the user's Microsoft 365 environment — emails, files, SharePoint pages, Teams messages, and people.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query text" },
        entityTypes: {
          type: "array",
          items: { type: "string" },
          description:
            "Content types to search: message, driveItem, site, listItem, drive, list, chatMessage, person. Default: all",
        },
        top: { type: "number", description: "Max results to return (default: 10, max: 25)" },
      },
      required: ["query"],
    },
  },
  {
    name: "configure_federated_knowledge",
    description:
      "Register a new federated knowledge configuration for Microsoft 365 Copilot. Requires admin rights and Copilot license.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string" },
        connectionId: { type: "string" },
      },
    },
  },
  {
    name: "delete_federated_knowledge",
    description: "Remove a federated knowledge configuration. Requires admin rights and Copilot license.",
    inputSchema: {
      type: "object",
      properties: {
        configurationId: { type: "string" },
      },
    },
  },
  {
    name: "ingest_federated_knowledge",
    description: "Trigger reingestion of a federated knowledge configuration. Requires admin rights and Copilot license.",
    inputSchema: {
      type: "object",
      properties: {
        configurationId: { type: "string" },
      },
    },
  },
  {
    name: "retrieve_federated_knowledge",
    description: "List all federated knowledge configurations. Requires admin rights and Copilot license.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
];

// ── Handler ───────────────────────────────────────────────────────────────────

export class KnowledgeGraphToolHandler {
  constructor(private token: string) {}

  async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    try {
      const text = await this.dispatch(name, args);
      const isError = /^(Error|Unknown)/.test(text);
      return { content: [{ type: "text", text }], ...(isError ? { isError: true } : {}) };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      log(`${name} error: ${msg}`);
      return { content: [{ type: "text", text: `Error: ${msg}` }], isError: true };
    }
  }

  private async dispatch(name: string, args: Record<string, unknown>): Promise<string> {
    switch (name) {
      case "query_federated_knowledge":
        return this.search(args);
      case "configure_federated_knowledge":
      case "delete_federated_knowledge":
      case "ingest_federated_knowledge":
      case "retrieve_federated_knowledge":
        return this.adminStub(name);
      default:
        return `Unknown knowledge tool: ${name}`;
    }
  }

  // ── M365 Search ─────────────────────────────────────────────────────────────

  private async search(args: Record<string, unknown>): Promise<string> {
    const query = String(args.query ?? "");
    const top = Math.min(Number(args.top ?? 10), 25);

    // Default entity types: search across everything
    const defaultTypes = ["message", "driveItem", "site", "listItem", "chatMessage"];
    const entityTypes =
      Array.isArray(args.entityTypes) && args.entityTypes.length > 0 ? args.entityTypes : defaultTypes;

    const res = await gf(this.token, "/search/query", {
      method: "POST",
      body: JSON.stringify({
        requests: [
          {
            entityTypes,
            query: { queryString: query },
            from: 0,
            size: top,
          },
        ],
      }),
    });

    if (!res.ok) return `Error searching M365 (${res.status}): ${JSON.stringify(res.body)}`;

    const data = res.body as any;
    const containers = data.value?.[0]?.hitsContainers ?? [];
    const allHits = containers.flatMap((c: any) => c.hits ?? []);
    const total = containers.reduce((sum: number, c: any) => sum + (c.total ?? 0), 0);

    if (allHits.length === 0) return `No results found for "${query}".`;

    const lines = [`Found ${total} result(s) for "${query}" (showing ${allHits.length}):\n`];
    for (const hit of allHits) {
      const r = hit.resource ?? {};
      const type = (r["@odata.type"] ?? "").replace("#microsoft.graph.", "");
      const title = r.subject ?? r.name ?? r.displayName ?? hit.hitId;
      const url = r.webUrl ?? "";
      const modified = r.lastModifiedDateTime ? new Date(r.lastModifiedDateTime).toLocaleDateString() : "";
      const from = r.from?.emailAddress
        ? ` | From: ${r.from.emailAddress.name ?? r.from.emailAddress.address}`
        : "";
      lines.push(
        [
          `[${type}] ${title}`,
          url ? `  URL: ${url}` : "",
          modified ? `  Modified: ${modified}${from}` : from ? `  ${from}` : "",
          hit.summary ? `  ${hit.summary.substring(0, 200)}` : "",
        ]
          .filter(Boolean)
          .join("\n")
      );
    }
    return lines.join("\n\n");
  }

  // ── Admin stubs ─────────────────────────────────────────────────────────────

  private adminStub(name: string): string {
    return [
      `"${name}" is a Microsoft 365 Copilot admin feature.`,
      "",
      "This tool manages federated knowledge connectors that link external data sources",
      "(e.g. ServiceNow, Confluence, SAP) into the Microsoft 365 Copilot experience.",
      "",
      "It requires:",
      "  • A Microsoft 365 Copilot Business Chat license",
      "  • Microsoft 365 admin or Search admin permissions",
      "",
      "If you need to search across M365 content (email, files, SharePoint, Teams),",
      'use "query_federated_knowledge" instead — it works with any M365 subscription.',
    ].join("\n");
  }
}
