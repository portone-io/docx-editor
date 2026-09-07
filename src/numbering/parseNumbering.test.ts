// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { fixtureNames, readFixture } from "../__testing__/docx";
import { readRunFormat } from "../docx/formatting";
import { importDocx } from "../docx/importDocx";
import { type Numbering, parseNumbering } from "./parseNumbering";

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
  it("resolves numbered list references with explicit plus signs", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="+7">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="+2"><w:abstractNumId w:val="+7"/></w:num>'
      )
    );
    expect(numbering.lists.get(2)?.levels.get(0)?.start).toBe(3);
  });

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
      restartAfterLevel: null,
      legal: false,
      suffix: "tab",
      align: "left",
      run: null,
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
      restartAfterLevel: null,
      legal: false,
      suffix: "tab",
      align: "left",
      run: null,
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
      restartAfterLevel: null,
      legal: false,
      suffix: "tab",
      align: "left",
      run: null,
    });
  });

  const formatted = (format: string) =>
    parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
          `<w:numFmt w:val="${format}"/><w:lvlText w:val="%1."/>` +
          "</w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    )
      .lists.get(1)
      ?.levels.get(0)?.format;

  it("a format no speller covers falls back to decimal and says so", () => {
    // A format the corpus does not use, so nothing spells it and the count is drawn as a number
    expect(formatted("japaneseCounting")).toBe("decimal");
    expect(formatted("")).toBe("decimal");
  });

  it("reads upperRoman, decimalZero, koreanDigital, ganada and chineseCounting", () => {
    for (const format of [
      "upperRoman",
      "decimalZero",
      "koreanDigital",
      "ganada",
      "chineseCounting",
    ]) {
      expect(formatted(format)).toBe(format);
    }
  });

  it("resolves numStyleLink through the numbering style's numPr to the linked abstractNum", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="Chapters"/></w:abstractNum>' +
          `<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Chapters"/>${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
          '<w:num w:numId="6"><w:abstractNumId w:val="1"/></w:num>'
      ),
      { links: new Map([["Chapters", 6]]) }
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.start).toBe(3);
  });

  it("follows numStyleLink to the definition standing behind that style when no style names a list", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="Chapters"/></w:abstractNum>' +
          `<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Chapters"/>${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.start).toBe(3);
  });

  it("a styleLink definition keeps its own levels rather than following the style back", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="0"><w:styleLink w:val="Chapters"/>${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
          '<w:lvlText w:val="\u25cf"/></w:lvl></w:abstractNum>' +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
          '<w:num w:numId="6"><w:abstractNumId w:val="1"/></w:num>'
      ),
      { links: new Map([["Chapters", 6]]) }
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.format).toBe("decimal");
  });

  it("leaves a numStyleLink that leads back to itself without levels rather than following it round", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:styleLink w:val="Chapters"/>' +
          '<w:numStyleLink w:val="Chapters"/></w:abstractNum>' +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      ),
      { links: new Map([["Chapters", 1]]) }
    );

    expect(numbering.lists.get(1)?.levels.size).toBe(0);
  });

  it("reads lvlRestart, isLgl, suff and lvlJc", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="2">' +
          '<w:numFmt w:val="decimal"/><w:lvlRestart w:val="1"/><w:isLgl/>' +
          '<w:suff w:val="space"/><w:lvlText w:val="%3."/><w:lvlJc w:val="end"/>' +
          "</w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(2)).toMatchObject({
      restartAfterLevel: 0,
      legal: true,
      suffix: "space",
      align: "right",
    });
  });

  it("a level that says none of them keeps what OOXML gives a level that says nothing", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="0">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(0)).toMatchObject({
      restartAfterLevel: null,
      legal: false,
      suffix: "tab",
      align: "left",
      run: null,
    });
  });

  it("reads a level that never restarts, and leaves an unreadable restart at the default", () => {
    const level = (restart: string) =>
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="1">' +
      `<w:numFmt w:val="decimal"/><w:lvlRestart w:val="${restart}"/>` +
      '<w:lvlText w:val="%2."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';

    expect(
      parseNumbering(numberingXml(level("0")))
        .lists.get(1)
        ?.levels.get(1)?.restartAfterLevel
    ).toBe(-1);
    expect(
      parseNumbering(numberingXml(level("every level")))
        .lists.get(1)
        ?.levels.get(1)?.restartAfterLevel
    ).toBeNull();
  });

  it("a suffix or a justification it does not know reads as the one OOXML gives by default", () => {
    const numbering = parseNumbering(
      numberingXml(
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
          '<w:numFmt w:val="decimal"/><w:suff w:val="comma"/>' +
          '<w:lvlText w:val="%1."/><w:lvlJc w:val="distribute"/>' +
          "</w:lvl></w:abstractNum>" +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      )
    );

    expect(numbering.lists.get(1)?.levels.get(0)).toMatchObject({
      suffix: "tab",
      align: "left",
      run: null,
    });
  });

  it("a level's rPr is read through readRun and is null without it", () => {
    const xml = numberingXml(
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
        '<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>' +
        "</w:lvl></w:abstractNum>" +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    );
    const levelOf = (numbering: Numbering) =>
      numbering.lists.get(1)?.levels.get(0)?.run;

    expect(
      levelOf(parseNumbering(xml, { readRun: (rPr) => readRunFormat(rPr) }))
    ).toEqual({ bold: true, color: "#FF0000" });
    expect(levelOf(parseNumbering(xml))).toBeNull();
  });

  it("a level that writes no rPr carries none even with a reader at hand", () => {
    const numbering = parseNumbering(
      numberingXml(
        `<w:abstractNum w:abstractNumId="0">${DECIMAL_LEVEL}</w:abstractNum>` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
      ),
      { readRun: (rPr) => readRunFormat(rPr) }
    );

    expect(numbering.lists.get(1)?.levels.get(0)?.run).toBeNull();
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
