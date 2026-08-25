/**
 * calendar-graph-tools.ts
 *
 * All calendar operations implemented directly via Microsoft Graph API.
 * Works with any standard M365 subscription — no Copilot license required.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [calendar-graph] ${msg}\n`);
}

interface GraphFetchResult {
  ok: boolean;
  status: number;
  body: any;
}

// ── Graph fetch helper ────────────────────────────────────────────────────────

async function gf(token: string, path: string, options: RequestInit = {}): Promise<GraphFetchResult> {
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="UTC"`,
      ...(options.headers ?? {}),
    },
  });
  const ct = res.headers.get("content-type") ?? "";
  const body = ct.includes("json") || ct.includes("odata")
    ? await res.json().catch(() => null)
    : await res.text().catch(() => "");
  return { ok: res.ok, status: res.status, body };
}

// ── Attendee helpers ──────────────────────────────────────────────────────────

interface Attendee {
  emailAddress: { name?: string; address: string };
  type: string;
}

function toAttendeeList(val: unknown): Attendee[] {
  if (!val) return [];
  const arr: unknown[] = Array.isArray(val)
    ? val
    : String(val)
        .split(/[,;]/)
        .map((s) => s.trim())
        .filter(Boolean);

  return arr
    .map((a): Attendee => {
      if (typeof a !== "string") return { emailAddress: { address: String(a) }, type: "required" };
      const m = a.match(/^(.+?)\s*<([^>]+)>$/);
      if (m) return { emailAddress: { name: m[1].trim(), address: m[2].trim() }, type: "required" };
      return { emailAddress: { address: a.trim() }, type: "required" };
    })
    .filter((a) => a.emailAddress.address.includes("@"));
}

function toForwardRecipients(val: unknown): Array<{ emailAddress: Attendee["emailAddress"] }> {
  return toAttendeeList(val).map((a) => ({ emailAddress: a.emailAddress }));
}

// ── Event formatter ───────────────────────────────────────────────────────────

function fmtEvent(ev: any): string {
  const start = ev.start;
  const end = ev.end;
  const loc = ev.location;
  const org = ev.organizer;
  const om = ev.onlineMeeting;
  const attendees = ev.attendees ?? [];
  const bodyText = (ev.body?.content ?? "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, 400);

  return [
    `ID: ${ev.id}`,
    `Subject: ${ev.subject ?? "(no subject)"}`,
    `Start: ${start?.dateTime ?? ""}${start?.timeZone ? ` (${start.timeZone})` : ""}`,
    `End: ${end?.dateTime ?? ""}`,
    loc?.displayName ? `Location: ${loc.displayName}` : "",
    org?.emailAddress ? `Organizer: ${org.emailAddress.name ?? org.emailAddress.address ?? ""}` : "",
    attendees.length > 0
      ? `Attendees:\n${attendees
          .map((a: any) => `  - ${a.emailAddress?.name ?? a.emailAddress?.address ?? ""} [${a.status?.response ?? "none"}]`)
          .join("\n")}`
      : "",
    ev.isOnlineMeeting ? `Online Meeting: Yes` : "",
    om?.joinUrl ? `Teams Join URL: ${om.joinUrl}` : "",
    bodyText ? `\nDescription:\n${bodyText}` : "",
  ]
    .filter(Boolean)
    .join("\n");
}

// ── Tool names set ────────────────────────────────────────────────────────────

export const CALENDAR_GRAPH_TOOL_NAMES = new Set([
  "ListEvents",
  "ListCalendarView",
  "CreateEvent",
  "UpdateEvent",
  "DeleteEventById",
  "AcceptEvent",
  "DeclineEvent",
  "TentativelyAcceptEvent",
  "CancelEvent",
  "ForwardEvent",
  "FindMeetingTimes",
  "GetRooms",
  "GetUserDateAndTimeZoneSettings",
]);

// ── Tool definitions ──────────────────────────────────────────────────────────

export const CALENDAR_GRAPH_TOOLS: Tool[] = [
  {
    name: "ListEvents",
    description: "List calendar events. Returns upcoming events ordered by start time, optionally filtered by date range.",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max events to return (default: 10, max: 50)" },
        startDateTime: {
          type: "string",
          description: "ISO 8601 — return events starting after this datetime",
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 — return events ending before this datetime",
        },
        filter: { type: "string", description: "OData $filter expression" },
        orderby: {
          type: "string",
          description: "Sort order (default: start/dateTime asc)",
        },
      },
    },
  },
  {
    name: "ListCalendarView",
    description: "Get all events (including recurring instances) within a specific date/time window.",
    inputSchema: {
      type: "object",
      properties: {
        startDateTime: {
          type: "string",
          description: "ISO 8601 start of the time range (required)",
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 end of the time range (required)",
        },
        top: { type: "number", description: "Max events to return (default: 20)" },
      },
      required: ["startDateTime", "endDateTime"],
    },
  },
  {
    name: "CreateEvent",
    description: "Create a new calendar event, optionally with attendees and a Teams meeting link.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Event title" },
        body: { type: "string", description: "Event description" },
        contentType: { type: "string", description: "'HTML' or 'text' (default: text)" },
        startDateTime: { type: "string", description: "ISO 8601 start date-time" },
        endDateTime: { type: "string", description: "ISO 8601 end date-time" },
        timeZone: {
          type: "string",
          description: "IANA time zone (e.g. Europe/Athens). Default: UTC",
        },
        location: { type: "string", description: "Meeting room or address" },
        attendees: {
          description: "Attendee email addresses",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        isOnlineMeeting: {
          type: "boolean",
          description: "Create a Teams online meeting link",
        },
        isAllDay: { type: "boolean", description: "Mark as an all-day event" },
      },
      required: ["subject", "startDateTime", "endDateTime"],
    },
  },
  {
    name: "UpdateEvent",
    description: "Update an existing calendar event.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "ID of the event to update" },
        subject: { type: "string" },
        body: { type: "string" },
        contentType: { type: "string" },
        startDateTime: { type: "string" },
        endDateTime: { type: "string" },
        timeZone: { type: "string" },
        location: { type: "string" },
        attendees: {
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        isOnlineMeeting: { type: "boolean" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "DeleteEventById",
    description: "Delete a calendar event by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "ID of the event to delete" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "AcceptEvent",
    description: "Accept a calendar event invitation.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string" },
        comment: { type: "string", description: "Optional acceptance message" },
        sendResponse: {
          type: "boolean",
          description: "Send response to organizer (default: true)",
        },
      },
      required: ["eventId"],
    },
  },
  {
    name: "DeclineEvent",
    description: "Decline a calendar event invitation.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string" },
        comment: { type: "string", description: "Optional decline message" },
        sendResponse: { type: "boolean", description: "Send response to organizer (default: true)" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "TentativelyAcceptEvent",
    description: "Tentatively accept a calendar event invitation.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string" },
        comment: { type: "string" },
        sendResponse: { type: "boolean" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "CancelEvent",
    description: "Cancel a calendar event and notify all attendees (organizer only).",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string" },
        comment: { type: "string", description: "Cancellation message sent to attendees" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "ForwardEvent",
    description: "Forward a calendar event invitation to additional recipients.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string" },
        to: {
          description: "Recipients to forward to",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        comment: { type: "string", description: "Optional message" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "FindMeetingTimes",
    description: "Find available meeting time slots for a group of attendees.",
    inputSchema: {
      type: "object",
      properties: {
        attendees: {
          description: "Attendee email addresses",
          oneOf: [{ type: "string" }, { type: "array", items: { type: "string" } }],
        },
        durationMinutes: {
          type: "number",
          description: "Required meeting duration in minutes (default: 60)",
        },
        startDateTime: {
          type: "string",
          description: "ISO 8601 — search from this datetime (default: now)",
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 — search until this datetime (default: one week from now)",
        },
        maxCandidates: {
          type: "number",
          description: "Max time slot suggestions to return (default: 5)",
        },
      },
      required: ["attendees"],
    },
  },
  {
    name: "GetRooms",
    description: "List available meeting rooms in the organization.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
  {
    name: "GetUserDateAndTimeZoneSettings",
    description: "Get the signed-in user's mailbox timezone, date format, and language settings.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
];

// ── Handler class ─────────────────────────────────────────────────────────────

export class CalendarGraphToolHandler {
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
      case "ListEvents":
        return this.listEvents(args);
      case "ListCalendarView":
        return this.calendarView(args);
      case "CreateEvent":
        return this.createEvent(args);
      case "UpdateEvent":
        return this.updateEvent(args);
      case "DeleteEventById":
        return this.deleteEvent(args);
      case "AcceptEvent":
        return this.respond(args, "accept");
      case "DeclineEvent":
        return this.respond(args, "decline");
      case "TentativelyAcceptEvent":
        return this.respond(args, "tentativelyAccept");
      case "CancelEvent":
        return this.cancelEvent(args);
      case "ForwardEvent":
        return this.forwardEvent(args);
      case "FindMeetingTimes":
        return this.findMeetingTimes(args);
      case "GetRooms":
        return this.getRooms();
      case "GetUserDateAndTimeZoneSettings":
        return this.getMailboxSettings();
      default:
        return `Unknown calendar tool: ${name}`;
    }
  }

  // ── List / View ─────────────────────────────────────────────────────────────

  private async listEvents(args: Record<string, unknown>): Promise<string> {
    const top = Math.min(Number(args.top ?? 10), 50);
    const params = new URLSearchParams({
      $top: String(top),
      $orderby: String(args.orderby ?? "start/dateTime asc"),
      $select: "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,body",
    });
    if (args.filter) {
      params.set("$filter", String(args.filter));
    } else if (args.startDateTime || args.endDateTime) {
      const filters: string[] = [];
      if (args.startDateTime) filters.push(`start/dateTime ge '${args.startDateTime}'`);
      if (args.endDateTime) filters.push(`end/dateTime le '${args.endDateTime}'`);
      params.set("$filter", filters.join(" and "));
    }

    const res = await gf(this.token, `/me/events?${params}`);
    if (!res.ok) return `Error listing events (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body;
    const events = data.value ?? [];
    if (events.length === 0) return "No events found.";
    return events.map((e: any) => fmtEvent(e)).join("\n\n---\n\n");
  }

  private async calendarView(args: Record<string, unknown>): Promise<string> {
    const top = Math.min(Number(args.top ?? 20), 50);
    const params = new URLSearchParams({
      startDateTime: String(args.startDateTime ?? ""),
      endDateTime: String(args.endDateTime ?? ""),
      $top: String(top),
      $select: "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,body",
    });
    const res = await gf(this.token, `/me/calendarView?${params}`);
    if (!res.ok) return `Error getting calendar view (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body;
    const events = data.value ?? [];
    if (events.length === 0) return "No events found in this time range.";
    return events.map((e: any) => fmtEvent(e)).join("\n\n---\n\n");
  }

  // ── Create / Update / Delete ─────────────────────────────────────────────────

  private async createEvent(args: Record<string, unknown>): Promise<string> {
    const tz = String(args.timeZone ?? "UTC");
    const ct = String(args.contentType ?? "text");
    const payload: Record<string, unknown> = {
      subject: String(args.subject ?? ""),
      body: {
        contentType: ct.toLowerCase() === "html" ? "HTML" : "Text",
        content: String(args.body ?? args.description ?? ""),
      },
      start: { dateTime: String(args.startDateTime ?? ""), timeZone: tz },
      end: { dateTime: String(args.endDateTime ?? ""), timeZone: tz },
    };
    if (args.location) payload.location = { displayName: String(args.location) };
    if (args.attendees) payload.attendees = toAttendeeList(args.attendees);
    if (args.isOnlineMeeting !== undefined) payload.isOnlineMeeting = args.isOnlineMeeting;
    if (args.isAllDay) payload.isAllDay = true;

    const res = await gf(this.token, "/me/events", { method: "POST", body: JSON.stringify(payload) });
    if (!res.ok) return `Error creating event (${res.status}): ${JSON.stringify(res.body)}`;
    const ev = res.body;
    return [
      "Event created.",
      `ID: ${ev.id}`,
      `Subject: ${ev.subject ?? ""}`,
      ev.onlineMeeting?.joinUrl ? `Teams Join URL: ${ev.onlineMeeting.joinUrl}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }

  private async updateEvent(args: Record<string, unknown>): Promise<string> {
    const id = String(args.eventId ?? args.id ?? "");
    const tz = String(args.timeZone ?? "UTC");
    const patch: Record<string, unknown> = {};
    if (args.subject !== undefined) patch.subject = args.subject;
    if (args.body !== undefined) {
      const ct = String(args.contentType ?? "text");
      patch.body = { contentType: ct.toLowerCase() === "html" ? "HTML" : "Text", content: String(args.body) };
    }
    if (args.startDateTime !== undefined) patch.start = { dateTime: String(args.startDateTime), timeZone: tz };
    if (args.endDateTime !== undefined) patch.end = { dateTime: String(args.endDateTime), timeZone: tz };
    if (args.location !== undefined) patch.location = { displayName: String(args.location) };
    if (args.attendees !== undefined) patch.attendees = toAttendeeList(args.attendees);
    if (args.isOnlineMeeting !== undefined) patch.isOnlineMeeting = args.isOnlineMeeting;

    const res = await gf(this.token, `/me/events/${id}`, { method: "PATCH", body: JSON.stringify(patch) });
    if (!res.ok) return `Error updating event (${res.status}): ${JSON.stringify(res.body)}`;
    return "Event updated.";
  }

  private async deleteEvent(args: Record<string, unknown>): Promise<string> {
    const id = String(args.eventId ?? args.id ?? "");
    const res = await gf(this.token, `/me/events/${id}`, { method: "DELETE" });
    if (!res.ok) return `Error deleting event (${res.status}): ${JSON.stringify(res.body)}`;
    return "Event deleted.";
  }

  // ── RSVP ────────────────────────────────────────────────────────────────────

  private async respond(
    args: Record<string, unknown>,
    action: "accept" | "decline" | "tentativelyAccept"
  ): Promise<string> {
    const id = String(args.eventId ?? args.id ?? "");
    const res = await gf(this.token, `/me/events/${id}/${action}`, {
      method: "POST",
      body: JSON.stringify({
        comment: String(args.comment ?? ""),
        sendResponse: args.sendResponse !== false,
      }),
    });
    if (!res.ok) return `Error responding to event (${res.status}): ${JSON.stringify(res.body)}`;
    const label = action === "accept" ? "accepted" : action === "decline" ? "declined" : "tentatively accepted";
    return `Event ${label}.`;
  }

  private async cancelEvent(args: Record<string, unknown>): Promise<string> {
    const id = String(args.eventId ?? args.id ?? "");
    const res = await gf(this.token, `/me/events/${id}/cancel`, {
      method: "POST",
      body: JSON.stringify({ comment: String(args.comment ?? "") }),
    });
    if (!res.ok) return `Error canceling event (${res.status}): ${JSON.stringify(res.body)}`;
    return "Event canceled and attendees notified.";
  }

  private async forwardEvent(args: Record<string, unknown>): Promise<string> {
    const id = String(args.eventId ?? args.id ?? "");
    const res = await gf(this.token, `/me/events/${id}/forward`, {
      method: "POST",
      body: JSON.stringify({
        comment: String(args.comment ?? ""),
        toRecipients: toForwardRecipients(args.to ?? args.toRecipients),
      }),
    });
    if (!res.ok) return `Error forwarding event (${res.status}): ${JSON.stringify(res.body)}`;
    return "Event forwarded.";
  }

  // ── Find meeting times ───────────────────────────────────────────────────────

  private async findMeetingTimes(args: Record<string, unknown>): Promise<string> {
    const now = new Date().toISOString();
    const oneWeek = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    const mins = Number(args.durationMinutes ?? 60);
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    const duration = `PT${h > 0 ? h + "H" : ""}${m > 0 ? m + "M" : ""}`;
    const attendees = toAttendeeList(args.attendees).map((a) => ({
      type: "required",
      emailAddress: a.emailAddress,
    }));

    const res = await gf(this.token, "/me/findMeetingTimes", {
      method: "POST",
      body: JSON.stringify({
        attendees,
        timeConstraint: {
          activityDomain: "work",
          timeslots: [
            {
              start: { dateTime: String(args.startDateTime ?? now), timeZone: "UTC" },
              end: { dateTime: String(args.endDateTime ?? oneWeek), timeZone: "UTC" },
            },
          ],
        },
        meetingDuration: duration,
        maxCandidates: Number(args.maxCandidates ?? 5),
        isOrganizerOptional: false,
        returnSuggestionReasons: true,
      }),
    });
    if (!res.ok) return `Error finding meeting times (${res.status}): ${JSON.stringify(res.body)}`;
    const data = res.body;
    const suggestions = data.meetingTimeSuggestions ?? [];
    if (suggestions.length === 0) return "No available meeting times found in the specified range.";
    return suggestions
      .map(
        (s: any, i: number) =>
          `Option ${i + 1}:\n  Start: ${s.meetingTimeSlot?.start?.dateTime ?? ""}\n  End:   ${
            s.meetingTimeSlot?.end?.dateTime ?? ""
          }\n  Confidence: ${Math.round((s.confidence ?? 0) * 100)}%`
      )
      .join("\n\n");
  }

  // ── Rooms / Mailbox settings ─────────────────────────────────────────────────

  private async getRooms(): Promise<string> {
    const res = await gf(this.token, "/places/microsoft.graph.room?$top=30");
    if (!res.ok) {
      // Fallback to legacy endpoint
      const fb = await gf(this.token, "/me/findRooms");
      if (!fb.ok) return "No rooms found or missing Places.Read.All permission.";
      const d = fb.body;
      const rooms = d.value ?? [];
      if (rooms.length === 0) return "No meeting rooms found.";
      return rooms.map((r: any) => `${r.name ?? ""} — ${r.address ?? ""}`).join("\n");
    }
    const data = res.body;
    const rooms = data.value ?? [];
    if (rooms.length === 0) return "No meeting rooms found.";
    return rooms
      .map((r: any) =>
        [
          `Name: ${r.displayName ?? ""}`,
          `Email: ${r.emailAddress ?? ""}`,
          r.building ? `Building: ${r.building}` : "",
          r.capacity ? `Capacity: ${r.capacity}` : "",
        ]
          .filter(Boolean)
          .join(" | ")
      )
      .join("\n");
  }

  private async getMailboxSettings(): Promise<string> {
    const res = await gf(this.token, "/me/mailboxSettings");
    if (!res.ok) return `Error getting mailbox settings (${res.status}): ${JSON.stringify(res.body)}`;
    const s = res.body;
    return [
      `Time Zone: ${s.timeZone ?? "(not set)"}`,
      `Language: ${s.language?.displayName ?? ""} (${s.language?.locale ?? ""})`,
      s.dateFormat ? `Date Format: ${s.dateFormat}` : "",
      s.timeFormat ? `Time Format: ${s.timeFormat}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }
}
