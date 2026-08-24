// @vitest-environment node
import { describe, expect, it } from "vitest";
import { DEFAULT_FONT_FALLBACKS, isEastAsianFontName } from "./fontStack";
import { DEFAULT_FONTS } from "./fonts";

/**
 * The built-in defaults are what a consumer gets with no props at all, and they
 * are deliberately all Latin: language-specific lists and fallbacks are the
 * consumer's to hand in.
 */
describe("the built-in defaults are Latin", () => {
  it("the built-in font list offers Latin faces only", () => {
    expect(DEFAULT_FONTS.filter(isEastAsianFontName)).toEqual([]);
  });

  it("the built-in default font is a Latin one", () => {
    expect(isEastAsianFontName(DEFAULT_FONT_FALLBACKS.defaultFontName)).toBe(
      false
    );
  });

  it("the built-in default font is the font its stack actually renders", () => {
    const [first] = DEFAULT_FONT_FALLBACKS.defaultStack.split(",");
    expect(first?.trim().replace(/^"|"$/g, "")).toBe(
      DEFAULT_FONT_FALLBACKS.defaultFontName
    );
  });
});
