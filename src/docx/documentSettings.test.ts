// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml } from "../ooxml/xml";
import { readCompatSettings, readDefaultTabStop } from "./documentSettings";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe("the document tab interval", () => {
  it("reads the setting from twips into points", () => {
    const settings = parseXml(
      `<w:settings ${W_NS}><w:defaultTabStop w:val="960"/></w:settings>`
    );
    expect(readDefaultTabStop(settings)).toBe(48);
  });

  it("leaves an absent or invalid setting to the OOXML fallback", () => {
    expect(readDefaultTabStop(null)).toBeNull();
    expect(
      readDefaultTabStop(
        parseXml(
          `<w:settings ${W_NS}><w:defaultTabStop w:val="0"/></w:settings>`
        )
      )
    ).toBeNull();
  });
});

describe("the compatibility settings", () => {
  function compat(children: string) {
    return readCompatSettings(
      parseXml(
        `<w:settings ${W_NS}><w:compat>${children}</w:compat></w:settings>`
      )
    );
  }

  it("reads noTabHangInd as switched on when present, and honors an explicit off", () => {
    expect(compat("<w:noTabHangInd/>")).toEqual({ noTabHangInd: true });
    expect(compat('<w:noTabHangInd w:val="0"/>')).toEqual({
      noTabHangInd: false,
    });
    expect(compat("<w:noSpaceRaiseLower/>")).toEqual({ noTabHangInd: false });
  });

  it("is switched off without a settings part or a compat element", () => {
    expect(readCompatSettings(null)).toEqual({ noTabHangInd: false });
    expect(
      readCompatSettings(parseXml(`<w:settings ${W_NS}></w:settings>`))
    ).toEqual({ noTabHangInd: false });
  });
});
