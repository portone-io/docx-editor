import { beforeAll, describe, expect, it } from "vitest";
import { commentTime, shownCommentDate } from "./commentDate";

// `vitest.config.ts` runs the suite in Seoul, nine hours ahead of UTC
beforeAll(() => {
  expect(new Date(Date.UTC(2026, 8, 30)).getTimezoneOffset()).toBe(-540);
});

/** The instant as a reader on this zone's clock sees it, in American English */
const onClockOf = (timeZone: string): string =>
  new Intl.DateTimeFormat("en-US", {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone,
  }).format(Date.UTC(2026, 8, 30, 6));

describe("shownCommentDate", () => {
  it("shows the instant on the reader's clock, with no zone named", () => {
    const shown = shownCommentDate("2026-09-30T06:00:00.000Z", "en-US");

    expect(shown).toBe(onClockOf("Asia/Seoul"));
    expect(shown).not.toBe(onClockOf("UTC"));
    expect(shown).toContain("Sep 30, 2026");
    expect(shown).toContain("3:00");
    expect(shown).not.toContain("UTC");
  });

  it("reads the browser's language where none is named", () => {
    expect(shownCommentDate("2026-09-30T06:00:00.000Z")).toBe(
      new Intl.DateTimeFormat(undefined, {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(Date.UTC(2026, 8, 30, 6))
    );
  });

  it("shows a date naming no instant as the document wrote it", () => {
    expect(shownCommentDate("someday", "en-US")).toBe("someday");
    // The browser would read this as the second of March
    expect(shownCommentDate("2026-02-30T10:00:00Z", "en-US")).toBe(
      "2026-02-30T10:00:00Z"
    );
    expect(shownCommentDate(null)).toBeNull();
  });
});

describe("commentTime", () => {
  it("answers the instant a date names, and null for one naming none", () => {
    expect(commentTime("2026-09-30T06:00:00.000Z")).toBe(
      Date.UTC(2026, 8, 30, 6)
    );
    expect(commentTime("2026-09-30T06:00:00Z")).toBeNull();
    expect(commentTime("2026-02-30T10:00:00Z")).toBeNull();
    expect(commentTime(null)).toBeNull();
  });
});
