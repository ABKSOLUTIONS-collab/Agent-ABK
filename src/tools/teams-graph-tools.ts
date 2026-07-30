/**
 * teams-graph-tools.ts
 *
 * Teams meeting operations via Microsoft Graph API.
 * Works with any standard M365 subscription — no Copilot license required,
 * except GetOnlineMeetingAiInsights which is a Copilot-only feature.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [teams-graph] ${msg}\n`);
}

interface GraphFetchResult {
  ok: boolean;
  status: number;
  body: unknown;
}

// ── Graph fetch helper ────────────────────────────────────────────────────────

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

// ── Meeting resolver ──────────────────────────────────────────────────────────

async function getMeetingId(token: string, args: Record<string, unknown>): Promise<string> {
  // Direct meeting ID
  const directId = String(args.meetingId ?? args.onlineMeetingId ?? "");
  if (directId) return directId;

  // Resolve by join URL
  const joinUrl = String(args.joinUrl ?? args.joinWebUrl ?? "");
  if (!joinUrl) throw new Error("Provide meetingId or joinUrl to identify the Teams meeting.");

  const encoded = encodeURIComponent(joinUrl);
  const res = await gf(token, `/me/onlineMeetings?$filter=joinWebUrl eq '${encoded}'`);
  if (!res.ok) throw new Error(`Failed to look up meeting by join URL (${res.status}): ${JSON.stringify(res.body)}`);
  const data = res.body as any;
  const meeting = data.value?.[0];
  if (!meeting) throw new Error("No meeting found for the given join URL.");
  return meeting.id;
}

// ── Tool names ────────────────────────────────────────────────────────────────

export const TEAMS_GRAPH_TOOL_NAMES = new Set([
  "GetOnlineMeetingTranscripts",
  "GetOnlineMeetingAttendanceReports",
  "GetOnlineMeetingAiInsights",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const TEAMS_GRAPH_TOOLS: Tool[] = [
  {
    name: "GetOnlineMeetingTranscripts",
    description: "Get the transcript(s) of a Teams online meeting. Provide the meeting ID or the Teams join URL.",
    inputSchema: {
      type: "object",
      properties: {
        meetingId: {
          type: "string",
          description: "Teams online meeting ID",
        },
        joinUrl: {
          type: "string",
          description: "Teams meeting join URL (e.g. https://teams.microsoft.com/l/meetup-join/...)",
        },
        transcriptId: {
          type: "string",
          description: "Specific transcript ID to retrieve. If omitted, lists all transcripts for the meeting.",
        },
      },
    },
  },
  {
    name: "GetOnlineMeetingAttendanceReports",
    description: "Get attendance report(s) for a Teams online meeting — who attended, join/leave times, duration.",
    inputSchema: {
      type: "object",
      properties: {
        meetingId: { type: "string", description: "Teams online meeting ID" },
        joinUrl: { type: "string", description: "Teams meeting join URL" },
        reportId: {
          type: "string",
          description: "Specific report ID. If omitted, returns the most recent report.",
        },
      },
    },
  },
  {
    name: "GetOnlineMeetingAiInsights",
    description: "Get Copilot-generated AI insights and summary for a Teams meeting. Requires Microsoft 365 Copilot license.",
    inputSchema: {
      type: "object",
      properties: {
        meetingId: { type: "string", description: "Teams online meeting ID" },
        joinUrl: { type: "string", description: "Teams meeting join URL" },
      },
    },
  },
];

// ── Handler ───────────────────────────────────────────────────────────────────

export class TeamsGraphToolHandler {
  constructor(private token: string) {}

  async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }> }> {
    try {
      const text = await this.dispatch(name, args);
      return { content: [{ type: "text", text }] };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      log(`${name} error: ${msg}`);
      return { content: [{ type: "text", text: `Error: ${msg}` }] };
    }
  }

  private async dispatch(name: string, args: Record<string, unknown>): Promise<string> {
    switch (name) {
      case "GetOnlineMeetingTranscripts":
        return this.getTranscripts(args);
      case "GetOnlineMeetingAttendanceReports":
        return this.getAttendance(args);
      case "GetOnlineMeetingAiInsights":
        return this.getAiInsights(args);
      default:
        return `Unknown Teams tool: ${name}`;
    }
  }

  // ── Transcripts ─────────────────────────────────────────────────────────────

  private async getTranscripts(args: Record<string, unknown>): Promise<string> {
    const meetingId = await getMeetingId(this.token, args);
    const specificId = args.transcriptId as string | undefined;

    if (specificId) {
      // Get content of a specific transcript
      const contentRes = await gf(this.token, `/me/onlineMeetings/${meetingId}/transcripts/${specificId}/content`, {
        headers: { Accept: "text/vtt, application/json" },
      });
      if (!contentRes.ok) return `Error getting transcript content (${contentRes.status}): ${JSON.stringify(contentRes.body)}`;
      return `Transcript content:\n\n${contentRes.body}`;
    }

    // List all transcripts
    const listRes = await gf(this.token, `/me/onlineMeetings/${meetingId}/transcripts`);
    if (!listRes.ok) return `Error listing transcripts (${listRes.status}): ${JSON.stringify(listRes.body)}`;
    const data = listRes.body as any;
    const transcripts = data.value ?? [];
    if (transcripts.length === 0)
      return "No transcripts found for this meeting. Transcripts are only available if recording was enabled.";

    if (transcripts.length === 1) {
      // Automatically fetch the single transcript's content
      const contentRes = await gf(
        this.token,
        `/me/onlineMeetings/${meetingId}/transcripts/${transcripts[0].id}/content`,
        { headers: { Accept: "text/vtt, application/json" } }
      );
      if (contentRes.ok) return `Transcript (${transcripts[0].createdDateTime ?? ""}):\n\n${contentRes.body}`;
    }

    return [
      `Found ${transcripts.length} transcript(s):`,
      ...transcripts.map(
        (t: any, i: number) => `${i + 1}. ID: ${t.id}\n   Created: ${t.createdDateTime ?? ""}`
      ),
      "\nUse transcriptId parameter to retrieve a specific transcript's content.",
    ].join("\n");
  }

  // ── Attendance ──────────────────────────────────────────────────────────────

  private async getAttendance(args: Record<string, unknown>): Promise<string> {
    const meetingId = await getMeetingId(this.token, args);
    const specificReportId = args.reportId as string | undefined;

    // List attendance reports
    const reportsRes = await gf(this.token, `/me/onlineMeetings/${meetingId}/attendanceReports`);
    if (!reportsRes.ok) return `Error getting attendance reports (${reportsRes.status}): ${JSON.stringify(reportsRes.body)}`;
    const reportsData = reportsRes.body as any;
    const reports = reportsData.value ?? [];
    if (reports.length === 0) return "No attendance reports found for this meeting.";

    const reportId = specificReportId ?? reports[0].id;
    const report = reports.find((r: any) => r.id === reportId) ?? reports[0];

    // Get attendance records for the report
    const recordsRes = await gf(this.token, `/me/onlineMeetings/${meetingId}/attendanceReports/${reportId}/attendanceRecords`);
    if (!recordsRes.ok) return `Error getting attendance records (${recordsRes.status}): ${JSON.stringify(recordsRes.body)}`;
    const recordsData = recordsRes.body as any;
    const records = recordsData.value ?? [];

    const lines = [
      `Attendance Report`,
      `Meeting Start: ${report.meetingStartDateTime ?? ""}`,
      `Meeting End:   ${report.meetingEndDateTime ?? ""}`,
      `Total Participants: ${report.totalParticipantCount ?? records.length}`,
      "",
      "Attendees:",
    ];
    for (const record of records) {
      const mins = Math.round((record.totalAttendanceInSeconds ?? 0) / 60);
      lines.push(`  ${record.displayName ?? record.emailAddress ?? record.id} — ${mins} min (${record.role ?? "attendee"})`);
      if (record.attendanceIntervals?.length) {
        for (const interval of record.attendanceIntervals) {
          lines.push(`    Joined: ${interval.joinDateTime ?? ""} | Left: ${interval.leaveDateTime ?? ""}`);
        }
      }
    }
    return lines.join("\n");
  }

  // ── AI Insights ─────────────────────────────────────────────────────────────

  private async getAiInsights(_args: Record<string, unknown>): Promise<string> {
    return [
      "Microsoft 365 Copilot AI Insights require a Copilot license.",
      "",
      "This feature generates AI-powered meeting summaries, action items, and insights.",
      "It is only available to users with a Microsoft 365 Copilot Business Chat subscription.",
      "",
      "Alternatives available without a Copilot license:",
      "• GetOnlineMeetingTranscripts — full transcript text of the meeting",
      "• GetOnlineMeetingAttendanceReports — who attended and for how long",
    ].join("\n");
  }
}
