// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readTabStopDirectives } from "./tabStops";
import { parseXml } from "./xml";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function pPr(inner: string): Element {
  return parseXml(`<w:pPr ${W_NS}>${inner}</w:pPr>`).documentElement;
}

describe("custom tab stop parsing", () => {
  it("normalizes transitional alignments and keeps future display data", () => {
    expect(
      readTabStopDirectives(
        pPr(
          '<w:tabs><w:tab w:val="left" w:pos="720" w:leader="dot"/>' +
            '<w:tab w:val="right" w:pos="1440"/>' +
            '<w:tab w:val="num" w:pos="2160"/>' +
            '<w:tab w:val="bar" w:pos="2880"/>' +
            '<w:tab w:val="clear" w:pos="3600"/></w:tabs>'
        )
      )
    ).toEqual([
      { positionPt: 36, align: "start", leader: "dot" },
      { positionPt: 72, align: "end" },
      { positionPt: 108, align: "num" },
      { positionPt: 144, align: "bar" },
      { positionPt: 180, align: "clear" },
    ]);
  });

  it("ignores entries that cannot be interpreted", () => {
    expect(
      readTabStopDirectives(
        pPr(
          '<w:tabs><w:tab w:val="start"/>' +
            '<w:tab w:val="unknown" w:pos="720"/></w:tabs>'
        )
      )
    ).toEqual([]);
  });
});
