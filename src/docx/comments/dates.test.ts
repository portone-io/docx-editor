import { beforeAll, describe, expect, it } from "vitest";
import {
  commentInstant,
  oneInstant,
  writtenCommentDates,
  writtenDate,
} from "./dates";

// `vitest.config.ts` runs the suite in Seoul, nine hours ahead of UTC with no daylight saving, so
// every expectation below says what the runtime's clock does there
beforeAll(() => {
  expect(new Date(Date.UTC(2026, 0, 1)).getTimezoneOffset()).toBe(-540);
  expect(new Date(Date.UTC(2026, 6, 1)).getTimezoneOffset()).toBe(-540);
});

describe("commentInstant", () => {
  it("takes dateUtc as the instant, whatever w:date says", () => {
    // What Word writes for a comment made at 10:32 in a zone seven hours behind UTC
    expect(commentInstant("2024-04-08T10:32:00Z", "2024-04-08T17:32:00Z")).toBe(
      "2024-04-08T17:32:00.000Z"
    );
  });

  it("reads w:date alone as the digits on the runtime's clock", () => {
    // 15:00 in Seoul is 06:00 UTC, which is what Word shows a reader in Seoul
    expect(commentInstant("2026-09-30T15:00:00Z", null)).toBe(
      "2026-09-30T06:00:00.000Z"
    );
  });

  it("sets aside any zone w:date names, as Word does", () => {
    expect(commentInstant("2026-09-30T15:00:00+02:00", null)).toBe(
      "2026-09-30T06:00:00.000Z"
    );
    expect(commentInstant("2026-09-30T15:00:00", null)).toBe(
      "2026-09-30T06:00:00.000Z"
    );
  });

  it("reads a w:date written without seconds or without a time", () => {
    expect(commentInstant("2026-09-30T15:00", null)).toBe(
      "2026-09-30T06:00:00.000Z"
    );
    expect(commentInstant("2026-09-30", null)).toBe("2026-09-29T15:00:00.000Z");
  });

  it("reads a dateUtc without a designator as UTC, and honours an offset it names", () => {
    expect(commentInstant(null, "2026-09-30T06:00:00")).toBe(
      "2026-09-30T06:00:00.000Z"
    );
    expect(commentInstant(null, "2026-09-30T15:00:00+09:00")).toBe(
      "2026-09-30T06:00:00.000Z"
    );
  });

  it("falls back to w:date where dateUtc is no date", () => {
    expect(commentInstant("2026-09-30T15:00:00Z", "yesterday")).toBe(
      "2026-09-30T06:00:00.000Z"
    );
  });

  it("hands back a w:date it cannot read as it was written", () => {
    expect(commentInstant("last Tuesday", null)).toBe("last Tuesday");
    expect(commentInstant("2026-02-30T10:00:00Z", null)).toBe(
      "2026-02-30T10:00:00Z"
    );
    expect(commentInstant("2026-09-30T25:00:00Z", null)).toBe(
      "2026-09-30T25:00:00Z"
    );
  });

  it("answers null for a comment that names no date", () => {
    expect(commentInstant(null, null)).toBeNull();
  });
});

describe("writtenCommentDates", () => {
  it("writes the wall clock into w:date and the instant into dateUtc", () => {
    expect(writtenCommentDates("2026-09-30T06:00:00.000Z")).toEqual({
      date: "2026-09-30T15:00:00Z",
      dateUtc: "2026-09-30T06:00:00Z",
    });
  });

  it("crosses the date line the wall clock crosses", () => {
    expect(writtenCommentDates("2026-09-30T20:30:15Z")).toEqual({
      date: "2026-10-01T05:30:15Z",
      dateUtc: "2026-09-30T20:30:15Z",
    });
  });

  it("keeps the instant to the second, the way Word writes them", () => {
    expect(writtenCommentDates("2026-09-30T06:00:00.999Z")).toEqual({
      date: "2026-09-30T15:00:00Z",
      dateUtc: "2026-09-30T06:00:00Z",
    });
  });

  it("reads back to the instant it was written for", () => {
    const { date, dateUtc } = writtenCommentDates("2026-09-30T06:00:00Z");
    expect(commentInstant(date, dateUtc)).toBe("2026-09-30T06:00:00.000Z");
    expect(commentInstant(date, null)).toBe("2026-09-30T06:00:00.000Z");
  });

  it("writes the current time when none is given", () => {
    const before = Math.floor(Date.now() / 1000) * 1000;
    const { date, dateUtc } = writtenCommentDates();
    const after = Date.now();
    const instant = Date.parse(commentInstant(date, dateUtc) ?? "");
    expect(instant).toBeGreaterThanOrEqual(before);
    expect(instant).toBeLessThanOrEqual(after);
    expect(writtenDate(date)).toBe(true);
    expect(writtenDate(dateUtc)).toBe(true);
  });

  it("puts a date that names no instant into w:date as given, recording no instant", () => {
    expect(writtenCommentDates("soon")).toEqual({
      date: "soon",
      dateUtc: null,
    });
    expect(writtenCommentDates("+010000-01-01T00:00:00.000Z")).toEqual({
      date: "+010000-01-01T00:00:00.000Z",
      dateUtc: null,
    });
  });
});

describe("writtenDate", () => {
  it("accepts the one form the writer puts out", () => {
    expect(writtenDate("2026-09-30T06:00:00Z")).toBe(true);
  });

  it("turns down every other spelling of a time, and a time the calendar lacks", () => {
    for (const value of [
      "2026-09-30T06:00:00.000Z",
      "2026-09-30T06:00:00",
      "2026-09-30T06:00:00+00:00",
      "2026-09-30T06:00Z",
      "2026-02-30T06:00:00Z",
      " 2026-09-30T06:00:00Z",
    ]) {
      expect(writtenDate(value), value).toBe(false);
    }
    expect(writtenDate(null)).toBe(false);
  });
});

describe("oneInstant", () => {
  it("accepts a wall clock up to fourteen hours either side of the instant", () => {
    expect(oneInstant("2026-09-30T15:00:00Z", "2026-09-30T06:00:00Z")).toBe(
      true
    );
    expect(oneInstant("2026-09-30T20:00:00Z", "2026-09-30T06:00:00Z")).toBe(
      true
    );
    expect(oneInstant("2026-09-29T20:00:00Z", "2026-09-30T06:00:00Z")).toBe(
      true
    );
  });

  it("turns down a pair no zone could have written", () => {
    expect(oneInstant("2026-09-30T20:00:01Z", "2026-09-30T06:00:00Z")).toBe(
      false
    );
    expect(oneInstant("2025-09-30T15:00:00Z", "2026-09-30T06:00:00Z")).toBe(
      false
    );
    expect(oneInstant("whenever", "2026-09-30T06:00:00Z")).toBe(false);
  });
});
