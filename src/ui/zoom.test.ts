import { describe, expect, it } from "vitest";
import { fitWidthZoom, normalizeZoom } from "./zoom";

describe("document zoom", () => {
  it("fits the page to the container without enlarging it", () => {
    expect(fitWidthZoom(1000, 794)).toBe(1);
    expect(fitWidthZoom(720, 794)).toBe(0.87);
    expect(fitWidthZoom(420, 794)).toBe(0.5);
  });

  it("keeps fixed zoom values within the supported range", () => {
    expect(normalizeZoom("fit-width")).toBe("fit-width");
    expect(normalizeZoom(0.1)).toBe(0.25);
    expect(normalizeZoom(1.25)).toBe(1.25);
    expect(normalizeZoom(3)).toBe(2);
    expect(normalizeZoom(Number.NaN)).toBe(1);
  });
});
