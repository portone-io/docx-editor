// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml } from "../ooxml/xml";
import { NO_THEME_FONTS, readThemeFonts, themeFontName } from "./theme";

const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** A theme part holding nothing but the font scheme, the one part of it we read */
function theme(fontScheme: string) {
  return readThemeFonts(
    parseXml(
      `<a:theme ${A_NS}><a:themeElements>${fontScheme}</a:themeElements></a:theme>`
    )
  );
}

const OFFICE_FONTS = theme(
  '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface="MS Gothic"/>' +
    '<a:cs typeface="Times New Roman"/>' +
    '<a:font script="Jpan" typeface="Meiryo"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="MS Mincho"/>' +
    '<a:cs typeface=""/></a:minorFont>' +
    "</a:fontScheme>"
);

describe("readThemeFonts", () => {
  it("reads the typeface of each script out of both schemes", () => {
    expect(OFFICE_FONTS).toEqual({
      major: {
        latin: "Cambria",
        eastAsia: "MS Gothic",
        complexScript: "Times New Roman",
      },
      minor: { latin: "Calibri", eastAsia: "MS Mincho", complexScript: null },
    });
  });

  it("a theme with no font scheme names no font at all", () => {
    expect(theme("")).toEqual(NO_THEME_FONTS);
    expect(theme('<a:fontScheme name="Office"/>')).toEqual(NO_THEME_FONTS);
  });
});

describe("themeFontName", () => {
  it("resolves every reference OOXML defines", () => {
    expect(themeFontName(OFFICE_FONTS, "majorHAnsi")).toBe("Cambria");
    expect(themeFontName(OFFICE_FONTS, "majorAscii")).toBe("Cambria");
    expect(themeFontName(OFFICE_FONTS, "majorEastAsia")).toBe("MS Gothic");
    expect(themeFontName(OFFICE_FONTS, "majorBidi")).toBe("Times New Roman");
    expect(themeFontName(OFFICE_FONTS, "minorHAnsi")).toBe("Calibri");
    expect(themeFontName(OFFICE_FONTS, "minorAscii")).toBe("Calibri");
    expect(themeFontName(OFFICE_FONTS, "minorEastAsia")).toBe("MS Mincho");
  });

  it("reads the reference regardless of how it is capitalized", () => {
    expect(themeFontName(OFFICE_FONTS, "minoreastasia")).toBe("MS Mincho");
    expect(themeFontName(OFFICE_FONTS, "MinorEastAsia")).toBe("MS Mincho");
  });

  it("a reference the theme leaves empty resolves to no font", () => {
    expect(themeFontName(OFFICE_FONTS, "minorBidi")).toBeNull();
    expect(themeFontName(NO_THEME_FONTS, "minorHAnsi")).toBeNull();
  });

  it("a reference we do not know resolves to no font", () => {
    expect(themeFontName(OFFICE_FONTS, "someTheme")).toBeNull();
    expect(themeFontName(OFFICE_FONTS, "")).toBeNull();
    expect(themeFontName(OFFICE_FONTS, null)).toBeNull();
  });
});
