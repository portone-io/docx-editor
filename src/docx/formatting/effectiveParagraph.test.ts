// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import type { Numbering } from "../../numbering/parseNumbering";
import { parseXml } from "../../ooxml/xml";
import {
  effectiveParagraphFormat,
  type ParagraphFormattingContext,
} from "./effectiveParagraph";
import { readStyles } from "./styles";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function context(
  defaults: ParagraphFormattingContext["defaults"],
  stylesXml: string,
  numbering: Numbering
): ParagraphFormattingContext {
  return {
    defaults,
    styles: readStyles(parseXml(`<w:styles ${W_NS}>${stylesXml}</w:styles>`)),
    defaultStyleId: null,
    numbering,
  };
}

const NUMBERING: Numbering = {
  lists: new Map([
    [
      4,
      {
        levels: new Map([
          [
            2,
            {
              format: "decimal",
              text: "%1.",
              start: 1,
              indent: null,
              tabStops: [
                { positionPt: 36, align: "clear" },
                { positionPt: 54, align: "num" },
              ],
            },
          ],
        ]),
      },
    ],
  ]),
};

describe("effective paragraph formatting", () => {
  it("layers defaults, numbering, style, and direct properties in OOXML order", () => {
    const formatting = context(
      {
        tabStops: [
          { positionPt: 36, align: "start" },
          { positionPt: 72, align: "start" },
        ],
      },
      '<w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:tabs>' +
        '<w:tab w:val="clear" w:pos="1440"/>' +
        '<w:tab w:val="center" w:pos="1800"/>' +
        "</w:tabs></w:pPr></w:style>",
      NUMBERING
    );

    expect(
      effectiveParagraphFormat(
        '<w:pPr><w:pStyle w:val="Body"/><w:numPr>' +
          '<w:ilvl w:val="2"/><w:numId w:val="4"/></w:numPr><w:tabs>' +
          '<w:tab w:val="end" w:pos="2160"/></w:tabs></w:pPr>',
        formatting
      )
    ).toEqual({
      numbering: { numId: 4, ilvl: 2 },
      tabStops: [
        { positionPt: 54, align: "num" },
        { positionPt: 90, align: "center" },
        { positionPt: 108, align: "end" },
      ],
    });
  });

  it("lets numId zero clear numbering inherited from a style", () => {
    const formatting = context(
      {},
      '<w:style w:type="paragraph" w:styleId="Listed"><w:pPr><w:numPr>' +
        '<w:ilvl w:val="2"/><w:numId w:val="4"/>' +
        "</w:numPr></w:pPr></w:style>",
      NUMBERING
    );

    expect(
      effectiveParagraphFormat(
        '<w:pPr><w:pStyle w:val="Listed"/><w:numPr>' +
          '<w:numId w:val="0"/></w:numPr></w:pPr>',
        formatting
      )
    ).toBeNull();
  });

  it("adds the implicit stop created by a hanging indent", () => {
    expect(
      effectiveParagraphFormat(
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>',
        context({}, "", { lists: new Map() })
      )
    ).toEqual({
      indentStartPt: 36,
      textIndentPt: -18,
      tabStops: [{ positionPt: 36, align: "start" }],
    });
    expect(
      effectiveParagraphFormat(
        '<w:pPr><w:bidi/><w:ind w:start="720" w:hanging="360"/></w:pPr>',
        context({}, "", { lists: new Map() })
      )
    ).toEqual({
      direction: "rtl",
      indentStartPt: 36,
      textIndentPt: -18,
      tabStops: [{ positionPt: 36, align: "start" }],
    });
  });

  it("layers Transitional and Strict leading indents in the same slot", () => {
    expect(
      effectiveParagraphFormat(
        '<w:pPr><w:ind w:left="720"/></w:pPr>',
        context({ indentStartPt: 18 }, "", { lists: new Map() })
      )
    ).toEqual({ indentStartPt: 36 });
  });
});
