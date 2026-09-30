/**
 * Word writes `w:date` as the author's wall clock with a `Z` it does not mean, and the instant as
 * `w16cex:dateUtc` (`spec/notes/comments.md`).
 */

/** `offsetMinutes` is null where no zone designator was written */
interface DateTimeFields {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  millisecond: number;
  offsetMinutes: number | null;
}

/** Lenient: the seconds, or the whole time, may be left out */
const DATE_TIME =
  /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?)?(Z|[+-]\d{2}:\d{2})?$/;

/** The form Word writes `w:date` and `dateUtc` in, and the only one this editor writes */
const WRITTEN = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/;

function offsetOf(zone: string | undefined): number | null {
  if (zone === undefined) return null;
  if (zone === "Z") return 0;
  const sign = zone.startsWith("-") ? -1 : 1;
  const [hours, minutes] = zone.slice(1).split(":").map(Number);
  return sign * (hours * 60 + minutes);
}

/** Rules out a 30 February or a 25:00 */
function onTheCalendar(fields: DateTimeFields): boolean {
  const date = new Date(0);
  date.setUTCFullYear(fields.year, fields.month - 1, fields.day);
  date.setUTCHours(fields.hour, fields.minute, fields.second);
  return (
    date.getUTCFullYear() === fields.year &&
    date.getUTCMonth() === fields.month - 1 &&
    date.getUTCDate() === fields.day &&
    date.getUTCHours() === fields.hour &&
    date.getUTCMinutes() === fields.minute &&
    date.getUTCSeconds() === fields.second
  );
}

function readDateTime(value: string): DateTimeFields | null {
  const match = DATE_TIME.exec(value);
  if (match === null) return null;
  const [, year, month, day, hour, minute, second, fraction, zone] = match;
  const fields: DateTimeFields = {
    year: Number(year),
    month: Number(month),
    day: Number(day),
    hour: Number(hour ?? 0),
    minute: Number(minute ?? 0),
    second: Number(second ?? 0),
    millisecond: Number((fraction ?? "").padEnd(3, "0").slice(0, 3)),
    offsetMinutes: offsetOf(zone),
  };
  const offset = fields.offsetMinutes;
  const zoneInRange = offset === null || Math.abs(offset) <= 14 * 60;
  return zoneInRange && onTheCalendar(fields) ? fields : null;
}

function utcInstant(fields: DateTimeFields): Date {
  const date = new Date(0);
  date.setUTCFullYear(fields.year, fields.month - 1, fields.day);
  date.setUTCHours(
    fields.hour,
    fields.minute - (fields.offsetMinutes ?? 0),
    fields.second,
    fields.millisecond
  );
  return date;
}

/** The digits read on the runtime's clock, any zone designator set aside, as Word reads `w:date` */
function floatingInstant(fields: DateTimeFields): Date {
  const date = new Date(0);
  date.setFullYear(fields.year, fields.month - 1, fields.day);
  date.setHours(fields.hour, fields.minute, fields.second, fields.millisecond);
  return date;
}

/**
 * When the comment was written, as an ISO 8601 instant: `dateUtc` where there is one, else `date`
 * on the runtime's clock. A `date` naming no calendar time comes back as written.
 */
export function commentInstant(
  date: string | null,
  dateUtc: string | null
): string | null {
  const utc = dateUtc === null ? null : readDateTime(dateUtc);
  if (utc !== null) return utcInstant(utc).toISOString();
  const floating = date === null ? null : readDateTime(date);
  return floating === null ? date : floatingInstant(floating).toISOString();
}

const pad = (value: number, width: number): string =>
  String(value).padStart(width, "0");

interface ClockReading {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

const runtimeClock = (instant: Date): ClockReading => ({
  year: instant.getFullYear(),
  month: instant.getMonth() + 1,
  day: instant.getDate(),
  hour: instant.getHours(),
  minute: instant.getMinutes(),
  second: instant.getSeconds(),
});

const utcClock = (instant: Date): ClockReading => ({
  year: instant.getUTCFullYear(),
  month: instant.getUTCMonth() + 1,
  day: instant.getUTCDate(),
  hour: instant.getUTCHours(),
  minute: instant.getUTCMinutes(),
  second: instant.getUTCSeconds(),
});

/** Null for a year outside 0001-9999, which the written form cannot hold */
function writtenForm(clock: ClockReading): string | null {
  const { year, month, day, hour, minute, second } = clock;
  if (year < 1 || year > 9999) return null;
  return `${pad(year, 4)}-${pad(month, 2)}-${pad(day, 2)}T${pad(hour, 2)}:${pad(minute, 2)}:${pad(second, 2)}Z`;
}

export interface WrittenCommentDates {
  /** `w:date`: the runtime's wall clock */
  date: string;
  dateUtc: string | null;
}

/**
 * The dates a comment written at `given`, or now, records, to the second as Word writes them. A
 * `given` naming no writable instant goes into `w:date` as given, with no `dateUtc`.
 */
export function writtenCommentDates(given?: string): WrittenCommentDates {
  const instant = given === undefined ? new Date() : new Date(given);
  const readable = !Number.isNaN(instant.valueOf());
  const date = readable ? writtenForm(runtimeClock(instant)) : null;
  const dateUtc = readable ? writtenForm(utcClock(instant)) : null;
  return date === null || dateUtc === null
    ? { date: given ?? instant.toISOString(), dateUtc: null }
    : { date, dateUtc };
}

export function writtenDate(value: string | null): boolean {
  return value !== null && WRITTEN.test(value) && readDateTime(value) !== null;
}

/** Whether some zone's clock could show `date` at `dateUtc`: `xsd:dateTime` offsets reach 14 hours */
export function oneInstant(date: string, dateUtc: string): boolean {
  const wall = readDateTime(date);
  const utc = readDateTime(dateUtc);
  if (wall === null || utc === null) return false;
  const apart =
    utcInstant({ ...wall, offsetMinutes: 0 }).valueOf() -
    utcInstant(utc).valueOf();
  return Math.abs(apart) <= 14 * 60 * 60 * 1000;
}
