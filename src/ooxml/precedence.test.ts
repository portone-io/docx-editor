import { describe, expect, it } from "vitest";
import type { XmlAttr } from "./element";
import { OVERRIDING_ATTRS, overridingAttrs, setAttr } from "./precedence";

describe("the attributes one attribute overrides", () => {
  it("drops the theme attributes when a font slot is written", () => {
    const rFonts: XmlAttr[] = [
      ["w:asciiTheme", "minorHAnsi"],
      ["w:eastAsiaTheme", "minorEastAsia"],
      ["w:hint", "eastAsia"],
    ];
    expect(setAttr(rFonts, "rFonts", "ascii", "Dotum")).toEqual([
      ["w:eastAsiaTheme", "minorEastAsia"],
      ["w:hint", "eastAsia"],
      ["w:ascii", "Dotum"],
    ]);
    // The slot that was not written keeps both its theme reference and its name
    expect(setAttr(rFonts, "rFonts", "ascii", "Dotum")).toContainEqual([
      "w:eastAsiaTheme",
      "minorEastAsia",
    ]);
    // The capitalization of the complex-script theme varies from document to document
    expect(overridingAttrs("rFonts", "cs")).toEqual(["cstheme", "csTheme"]);
  });

  it("drops the character-unit spelling when a twips indent is written", () => {
    const ind: XmlAttr[] = [
      ["w:leftChars", "200"],
      ["w:right", "100"],
    ];
    expect(setAttr(ind, "ind", "left", "720")).toEqual([
      ["w:right", "100"],
      ["w:left", "720"],
    ]);
    // Withdrawing the indent takes the character-unit spelling with it
    expect(setAttr([...ind, ["w:left", "720"]], "ind", "left", null)).toEqual([
      ["w:right", "100"],
    ]);
  });

  it("keeps shd's color apart from the color element's val", () => {
    const shd: XmlAttr[] = [
      ["w:val", "clear"],
      ["w:color", "auto"],
      ["w:themeFill", "accent1"],
    ];
    // `w:shd` records a color of its own, so writing its fill leaves that color standing
    expect(setAttr(shd, "shd", "fill", "FF0000")).toEqual([
      ["w:val", "clear"],
      ["w:color", "auto"],
      ["w:fill", "FF0000"],
    ]);
    // The `w:color` element records its color in `w:val` instead, and the theme color beats it
    expect(overridingAttrs("shd", "val")).toEqual([]);
    expect(overridingAttrs("color", "val")).toEqual([
      "themeColor",
      "themeTint",
      "themeShade",
    ]);
  });

  it("lists every OVERRIDING_ATTRS key exactly once", () => {
    const keys = Object.keys(OVERRIDING_ATTRS);
    expect(new Set(keys).size).toBe(keys.length);
    for (const [key, overrides] of Object.entries(OVERRIDING_ATTRS)) {
      const [element, name, ...rest] = key.split("/");
      expect(
        rest,
        `${key} names more than an element and an attribute`
      ).toEqual([]);
      expect(element.length, `${key} names no element`).toBeGreaterThan(0);
      expect(name.length, `${key} names no attribute`).toBeGreaterThan(0);
      expect(new Set(overrides).size, `${key} repeats an override`).toBe(
        overrides.length
      );
      // Two attributes that each override the other would leave the writer no way to write either
      for (const override of overrides) {
        expect(
          overridingAttrs(element, override),
          `${key} is overridden by what it overrides`
        ).toEqual([]);
      }
    }
  });
});
