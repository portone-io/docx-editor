/**
 * Null for a date that is not an instant as `toISOString` spells one: that is a date the document
 * wrote as no time, which the browser would still read, `2026-02-30` as the second of March.
 */
export function commentTime(value: string | null): number | null {
  if (value === null) return null;
  const time = Date.parse(value);
  return Number.isNaN(time) || new Date(time).toISOString() !== value
    ? null
    : time;
}

/** In the reader's zone and browser language; a date naming no instant is shown as written */
export function shownCommentDate(
  value: string | null,
  locales?: Intl.LocalesArgument
): string | null {
  const time = commentTime(value);
  return time === null
    ? value
    : new Intl.DateTimeFormat(locales, {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(time);
}
