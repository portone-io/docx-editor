// @vitest-environment node
import { describe, expect, it } from "vitest";
import { nextTabTarget, tabAdvancePt } from "./tabLayout";

describe("nextTabTarget", () => {
  it("skips bar stops and treats a legacy list stop as leading", () => {
    expect(
      nextTabTarget(
        20,
        [
          { positionPt: 24, align: "bar" },
          { positionPt: 48, align: "num" },
        ],
        36
      )
    ).toEqual({ align: "start", targetPt: 48 });
  });

  it("falls back to the next document interval", () => {
    expect(nextTabTarget(72, [], 36)).toEqual({
      align: "start",
      targetPt: 108,
    });
  });

  it("chooses the next stop independently of the following text width", () => {
    expect(nextTabTarget(30, [{ positionPt: 48, align: "end" }], 36)).toEqual({
      align: "end",
      targetPt: 48,
    });
  });
});

describe("tabAdvancePt", () => {
  const segment = { wholePt: 20, decimalPrefixPt: 8 };

  it.each([
    ["start", 26],
    ["center", 16],
    ["end", 6],
    ["decimal", 18],
  ] as const)("aligns a %s tab using the following segment", (align, width) => {
    expect(tabAdvancePt(10, { align, targetPt: 36 }, segment)).toBe(width);
  });

  it("does not move following text backward when an aligned stop overlaps it", () => {
    expect(tabAdvancePt(30, { align: "end", targetPt: 48 }, segment)).toBe(0);
  });
});
