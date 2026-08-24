// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml } from "../ooxml/xml";
import { readDefaultTabStop } from "./documentSettings";

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
