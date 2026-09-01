/**
 * format.ts
 *
 * Presentation of timestamps for anything a user reads.
 *
 * Microsoft Graph returns every timestamp in UTC, and the container has no TZ
 * set, so both raw Graph strings and bare toLocale* calls render as UTC. The
 * agent then repeats that verbatim: an email that Outlook shows at 11:07 was
 * being reported as 08:07, which reads as a plainly wrong answer rather than a
 * timezone detail — the user has Outlook open next to it.
 *
 * Every user-facing date therefore goes through here. The timezone is
 * configurable so this does not have to be edited if the organisation moves or
 * a second region is added, but it must never silently fall back to UTC.
 */

export const ORG_TIMEZONE = process.env.ORG_TIMEZONE || "Europe/Athens";

function toDate(v: unknown): Date | null {
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
  if (typeof v !== "string" || !v.trim()) return null;
  // Graph sometimes omits the zone designator on local-time fields (calendar
  // start/end come with a separate timeZone property). Those are UTC in every
  // call this bridge makes, so assume UTC rather than the container's locale,
  // which would shift them twice.
  const s = /[zZ]|[+-]\d{2}:\d{2}$/.test(v) ? v : `${v}Z`;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

const dateTimeFmt = new Intl.DateTimeFormat("en-GB", {
  timeZone: ORG_TIMEZONE,
  day: "2-digit", month: "2-digit", year: "numeric",
  hour: "2-digit", minute: "2-digit", hour12: false,
});

const dateFmt = new Intl.DateTimeFormat("en-GB", {
  timeZone: ORG_TIMEZONE,
  day: "2-digit", month: "2-digit", year: "numeric",
});

/**
 * "01/09/2026 11:07" in the organisation's timezone.
 * Returns the original value unchanged if it cannot be parsed, so a format
 * this does not recognise degrades to the raw string rather than to "Invalid
 * Date" or an empty field.
 */
export function fmtDateTime(v: unknown): string {
  const d = toDate(v);
  if (!d) return typeof v === "string" ? v : "";
  return dateTimeFmt.format(d).replace(",", "");
}

/** "01/09/2026" in the organisation's timezone. */
export function fmtDate(v: unknown): string {
  const d = toDate(v);
  if (!d) return typeof v === "string" ? v : "";
  return dateFmt.format(d);
}
