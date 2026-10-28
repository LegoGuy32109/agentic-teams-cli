/** A scheduled message must be at least this far ahead. */
export const MIN_LEAD_MS = 60_000;

const USAGE = "Use an ISO 8601 time with an offset (2026-10-20T09:30:00-04:00), HH:MM, or +<n>m or +<n>h.";

/**
 * Reads a `--at` value into a send time. Accepts ISO 8601 with `Z` or an offset, `HH:MM` (today in the
 * local zone, or tomorrow when that time has passed), and `+<n>m` or `+<n>h`.
 * Refuses a time in the past and a time less than one minute ahead.
 */
export function parseSendAt(input: string, now = new Date()): Date {
  const value = input.trim();
  let at: Date;
  let relative = false;
  let match: RegExpMatchArray | null;
  if ((match = value.match(/^\+(\d{1,4})([mh])$/))) {
    relative = true;
    at = new Date(now.getTime() + Number(match[1]) * (match[2] === "h" ? 3_600_000 : 60_000));
  } else if ((match = value.match(/^(\d{1,2}):(\d{2})$/))) {
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours > 23 || minutes > 59) throw new Error(`"${input}" is not a valid time. ${USAGE}`);
    at = new Date(now);
    at.setHours(hours, minutes, 0, 0);
    // Local-time arithmetic keeps the clock time across a daylight saving change.
    if (at.getTime() <= now.getTime()) at.setDate(at.getDate() + 1);
  } else if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/.test(value)) {
    at = new Date(value);
  } else {
    throw new Error(`"${input}" is not a valid --at time. ${USAGE}`);
  }
  if (Number.isNaN(at.getTime())) throw new Error(`"${input}" is not a valid time. ${USAGE}`);
  // Teams stores the time to the second.
  at = new Date(Math.floor(at.getTime() / 1000) * 1000);
  if (at.getTime() <= now.getTime()) throw new Error(`The send time ${at.toISOString()} is in the past.`);
  if (at.getTime() - now.getTime() < MIN_LEAD_MS) {
    throw new Error(`The send time must be at least 1 minute from now${relative ? "" : ` (${at.toISOString()} is too soon)`}.`);
  }
  return at;
}

/** Formats a time as local time with its zone, then UTC. */
export function describeSendAt(at: Date): { local: string; utc: string } {
  const local = at.toLocaleString("en-US", { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit", timeZoneName: "short" });
  return { local, utc: at.toISOString() };
}
