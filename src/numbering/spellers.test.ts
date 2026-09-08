import { describe, expect, it } from "vitest";
import { isNumberFormat, type NumberFormat, spellNumber } from "./spellers";

/** The cap the markers apply, which is what sends a spelling too long to grow back to a decimal */
const MAX_CHARS = 64;

function spelled(format: NumberFormat, counts: readonly number[]): string[] {
  return counts.map((count) => spellNumber(count, format, MAX_CHARS));
}

const FIRST_TEN = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

describe("the formats a marker is counted in", () => {
  it("decimal writes the count as it is", () => {
    expect(spelled("decimal", [1, 9, 10, 4000])).toEqual([
      "1",
      "9",
      "10",
      "4000",
    ]);
  });

  it("decimalZero puts a zero in front of the counts below ten", () => {
    expect(spelled("decimalZero", [1, 9, 10, 11, 100])).toEqual([
      "01",
      "09",
      "10",
      "11",
      "100",
    ]);
  });

  it("upperRoman and lowerRoman spell the numeral", () => {
    expect(spelled("upperRoman", [1, 4, 9, 14, 40, 1990])).toEqual([
      "I",
      "IV",
      "IX",
      "XIV",
      "XL",
      "MCMXC",
    ]);
    expect(spelled("lowerRoman", [1, 4, 9])).toEqual(["i", "iv", "ix"]);
  });

  it("the letters carry on the way Word does, the last one written twice over", () => {
    expect(spelled("upperLetter", [1, 26, 27, 54])).toEqual([
      "A",
      "Z",
      "AA",
      "BBB",
    ]);
    expect(spelled("lowerLetter", [1, 26, 27])).toEqual(["a", "z", "aa"]);
  });

  it("ganada counts through the fourteen Korean letters and then doubles them", () => {
    expect(spelled("ganada", [1, 2, 14, 15, 16, 29])).toEqual([
      "가",
      "나",
      "하",
      "가가",
      "나나",
      "가가가",
    ]);
  });

  it("koreanDigital writes the count digit by digit", () => {
    expect(spelled("koreanDigital", [1, 9, 10, 11, 20, 105])).toEqual([
      "일",
      "구",
      "일영",
      "일일",
      "이영",
      "일영오",
    ]);
  });

  /** The pattern §17.18.59 gives: 一, …, 十, 十一, …, 二十, …, 九十九, 一〇〇, 一〇一 */
  it("chineseCounting counts in tens up to ninety-nine and digit by digit past it", () => {
    expect(spelled("chineseCounting", FIRST_TEN)).toEqual([
      "一",
      "二",
      "三",
      "四",
      "五",
      "六",
      "七",
      "八",
      "九",
      "十",
    ]);
    expect(spelled("chineseCounting", [11, 19, 20, 21, 90, 99])).toEqual([
      "十一",
      "十九",
      "二十",
      "二十一",
      "九十",
      "九十九",
    ]);
    expect(spelled("chineseCounting", [100, 101, 110])).toEqual([
      "一〇〇",
      "一〇一",
      "一一〇",
    ]);
  });

  /**
   * A `%n` slot may name a bullet level, which has no symbol to put in the slot, so what goes in
   * is the count itself
   */
  it("a bullet level counted in a slot of another level's text reads as a number", () => {
    expect(spelled("bullet", [1, 12])).toEqual(["1", "12"]);
  });
});

/**
 * A crafted `w:start` can name a count no real list reaches, and a spelling that grows with the
 * count would then be drawn megabytes long in front of every paragraph of the list
 */
describe("a count far past what a list would use", () => {
  const bomb = 2_000_000_000;

  it.each([
    "lowerLetter",
    "upperLetter",
    "lowerRoman",
    "upperRoman",
    "ganada",
  ] as const)("draws %s as a decimal instead of spelling it out", (format) => {
    expect(spellNumber(bomb, format, MAX_CHARS)).toBe(`${bomb}`);
  });

  it("spells out the counts a list really uses", () => {
    expect(spellNumber(26 * MAX_CHARS, "upperLetter", MAX_CHARS)).toHaveLength(
      MAX_CHARS
    );
    expect(spellNumber(14 * MAX_CHARS, "ganada", MAX_CHARS)).toHaveLength(
      MAX_CHARS
    );
  });

  it("a format that writes the count digit by digit is drawn however far it counts", () => {
    expect(spellNumber(bomb, "koreanDigital", MAX_CHARS)).toBe(
      "이영영영영영영영영영"
    );
    expect(spellNumber(bomb, "decimal", MAX_CHARS)).toBe(`${bomb}`);
  });

  it("a count below one is drawn as the number it is", () => {
    expect(spellNumber(0, "upperLetter", MAX_CHARS)).toBe("0");
    expect(spellNumber(-3, "chineseCounting", MAX_CHARS)).toBe("-3");
  });
});

describe("isNumberFormat", () => {
  it("holds every format that has a speller", () => {
    expect(isNumberFormat("chineseCounting")).toBe(true);
    expect(isNumberFormat("bullet")).toBe(true);
  });

  it("turns down a format nothing spells, and one that is not written at all", () => {
    expect(isNumberFormat("japaneseCounting")).toBe(false);
    expect(isNumberFormat("hasOwnProperty")).toBe(false);
    expect(isNumberFormat(null)).toBe(false);
  });
});
