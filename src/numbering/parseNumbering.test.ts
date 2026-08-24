// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { fixtureNames, readFixture } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { parseNumbering } from "./parseNumbering";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function numberingXml(inner: string): string {
  return `<w:numbering ${W_NS}>${inner}</w:numbering>`;
}

const DECIMAL_LEVEL =
  '<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/>' +
  "</w:pPr></w:lvl>";

describe("parseNumbering", () => {
  it("there are no lists when there is no numbering.xml", () => {
    expect(parseNumbering(null).lists.size).toBe(0);
  });

  it("reads the levels of the abstractNum the num points at", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="7">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="2"><w:abstractNumId w:val="7"/></w:num>'
      )
    );
    expect(numbering.lists.get(2)?.levels.get(0)).toEqual({
      format: "decimal",
      text: "%1.",
      start: 3,
      indent: {
        startTwips: 720,
        endTwips: null,
        hangingTwips: 360,
        firstLineTwips: null,
      },
    });
  });

  it("the levels are empty when the abstractNum cannot be found", () => {
    const numbering = parseNumbering(
      numberingXml('<w:num w:numId="1"><w:abstractNumId w:val="9"/></w:num>')
    );
    expect(numbering.lists.get(1)?.levels.size).toBe(0);
  });

  it("startOverride changes only the start number", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="0">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/>' +
          '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/>' +
          "</w:lvlOverride></w:num>"
      )
    );
    expect(numbering.lists.get(1)?.levels.get(0)).toEqual({
      format: "decimal",
      text: "%1.",
      start: 10,
      indent: {
        startTwips: 720,
        endTwips: null,
        hangingTwips: 360,
        firstLineTwips: null,
      },
    });
  });

  it("an lvl inside lvlOverride swaps out the whole level", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="0">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/>' +
          '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">' +
          '<w:numFmt w:val="bullet"/><w:lvlText w:val="●"/>' +
          "</w:lvl></w:lvlOverride></w:num>"
      )
    );
    expect(numbering.lists.get(1)?.levels.get(0)).toEqual({
      format: "bullet",
      text: "●",
      start: 1,
      indent: null,
    });
  });

  it("a number format we do not know is left as decimal", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
          '<w:numFmt w:val="ganada"/><w:lvlText w:val="%1."/>' +
          "</w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );
    expect(numbering.lists.get(1)?.levels.get(0)?.format).toBe("decimal");
  });

  it("reads tab directives from numbering-level paragraph properties", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
          '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr>' +
          '<w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs>' +
          "</w:pPr></w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.tabStops).toEqual([
      { positionPt: 36, align: "num" },
    ]);
  });

  it("keeps Strict numbering indents on their logical sides", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
          '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr>' +
          '<w:ind w:start="720" w:end="240" w:hanging="360"/>' +
          "</w:pPr></w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.indent).toEqual({
      startTwips: 720,
      endTwips: 240,
      hangingTwips: 360,
      firstLineTwips: null,
    });
  });
});

describe("numbering in the fixtures", () => {
  it.each(fixtureNames)(
    "%s: carries numbering.xml into the session and reads it",
    (name) => {
      const { session } = importDocx(readFixture(name));
      expect(session.numberingXml).not.toBeNull();
      expect(parseNumbering(session.numberingXml).lists.size).toBeGreaterThan(
        0
      );
    }
  );
});
