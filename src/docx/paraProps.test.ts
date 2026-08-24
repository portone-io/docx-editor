// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import type { LineSpacing, ParagraphAlign } from "../model/format";
import { templateIndent } from "../numbering/listTemplate";
import { parseXml, W_NS } from "../ooxml/xml";
import { readStyles, type StyleTable } from "./formatting";
import {
  type ListChange,
  withLeftIndent,
  withLineSpacing,
  withListNumbering,
  withParagraphAlign,
  withParagraphStyle,
} from "./paraProps";

const NUM_PR = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>';

/** Turns it into a first-level list while applying that level's indents */
function toList(numId: number, ilvl = 0): ListChange {
  return {
    numbering: { numId, ilvl },
    indent: { kind: "level", indent: templateIndent(ilvl) },
  };
}

const LEAVE_LIST: ListChange = {
  numbering: null,
  indent: { kind: "clearHanging" },
};

function pPrOf(pPr: string | null, change: ListChange): string | null {
  const props = withListNumbering(pPr, change);
  if (!props) throw new Error("the fragment was not recognized");
  return props.pPr;
}

describe("swapping numPr in", () => {
  it("creates a fresh fragment for a paragraph that had no formatting", () => {
    expect(
      pPrOf(null, {
        numbering: { numId: 4, ilvl: 0 },
        indent: { kind: "keep" },
      })
    ).toBe(`<w:pPr>${NUM_PR}</w:pPr>`);
  });

  it("goes into its prescribed slot (after widowControl, before pBdr)", () => {
    const pPr =
      '<w:pPr><w:widowControl w:val="0"/>' +
      '<w:pBdr><w:top w:val="nil"/></w:pBdr>' +
      '<w:jc w:val="both"/></w:pPr>';
    expect(
      pPrOf(pPr, { numbering: { numId: 4, ilvl: 0 }, indent: { kind: "keep" } })
    ).toBe(
      '<w:pPr><w:widowControl w:val="0"/>' +
        NUM_PR +
        '<w:pBdr><w:top w:val="nil"/></w:pBdr>' +
        '<w:jc w:val="both"/></w:pPr>'
    );
  });

  it("an existing numPr keeps its slot and only its values change", () => {
    const pPr = `<w:pPr>${NUM_PR}<w:jc w:val="both"/></w:pPr>`;
    expect(
      pPrOf(pPr, { numbering: { numId: 7, ilvl: 2 }, indent: { kind: "keep" } })
    ).toBe(
      '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr>' +
        '<w:jc w:val="both"/></w:pPr>'
    );
  });

  it("formatting we do not read stays character for character", () => {
    const rest =
      '<w:spacing w:line="276" w:lineRule="auto"/>' +
      '<w:jc w:val="both"/>' +
      '<w:rPr><w:rFonts w:ascii="Malgun Gothic"/><w:sz w:val="20"/></w:rPr>';
    const pPr = `<w:pPr>${NUM_PR}${rest}</w:pPr>`;
    const change: ListChange = {
      numbering: { numId: 7, ilvl: 1 },
      indent: { kind: "keep" },
    };
    expect(pPrOf(pPr, change)).toBe(
      '<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>' +
        `${rest}</w:pPr>`
    );
  });

  it("leaves the opening tag's attributes alone", () => {
    const pPr = '<w:pPr w14:paraId="1F2A"><w:jc w:val="both"/></w:pPr>';
    expect(
      pPrOf(pPr, { numbering: { numId: 4, ilvl: 0 }, indent: { kind: "keep" } })
    ).toBe(`<w:pPr w14:paraId="1F2A">${NUM_PR}<w:jc w:val="both"/></w:pPr>`);
  });
});

describe("taking a paragraph out of a list", () => {
  it("removes numPr and clears only the hanging indent", () => {
    const pPr = `<w:pPr>${NUM_PR}<w:ind w:left="720" w:right="200" w:hanging="360"/></w:pPr>`;
    expect(pPrOf(pPr, LEAVE_LIST)).toBe(
      '<w:pPr><w:ind w:left="720" w:right="200"/></w:pPr>'
    );
  });

  it("a first-line indent is not a marker slot, so it stays", () => {
    const pPr = `<w:pPr>${NUM_PR}<w:ind w:left="400" w:firstLine="200"/></w:pPr>`;
    expect(pPrOf(pPr, LEAVE_LIST)).toBe(
      '<w:pPr><w:ind w:left="400" w:firstLine="200"/></w:pPr>'
    );
  });

  it("a fragment that held nothing but the list disappears entirely", () => {
    expect(pPrOf(`<w:pPr>${NUM_PR}</w:pPr>`, LEAVE_LIST)).toBeNull();
  });
});

describe("the indent the level sets", () => {
  it("changes the paragraph indent to that level's value", () => {
    const pPr = `<w:pPr>${NUM_PR}<w:ind w:left="1440" w:hanging="400"/></w:pPr>`;
    expect(pPrOf(pPr, toList(4, 1))).toBe(
      '<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="4"/></w:numPr>' +
        '<w:ind w:left="1440" w:hanging="360"/></w:pPr>'
    );
  });

  it("keeps the margins we do not set", () => {
    const pPr =
      '<w:pPr><w:ind w:start="1440" w:right="200" w:firstLine="200"/></w:pPr>';
    expect(pPrOf(pPr, toList(4))).toBe(
      `<w:pPr>${NUM_PR}<w:ind w:left="720" w:hanging="360" w:right="200"/></w:pPr>`
    );
  });

  it("keeps a direct trailing indent when the list level also defines one", () => {
    const indent = { ...templateIndent(0), endTwips: 480 };
    const change: ListChange = {
      numbering: { numId: 4, ilvl: 0 },
      indent: { kind: "level", indent },
    };

    expect(
      pPrOf('<w:pPr><w:ind w:start="360" w:end="200"/></w:pPr>', change)
    ).toBe(
      `<w:pPr>${NUM_PR}<w:ind w:left="720" w:hanging="360" w:end="200"/></w:pPr>`
    );
  });

  it("inserts it in order even into a paragraph that had no indent", () => {
    const pPr = '<w:pPr><w:spacing w:after="240"/><w:jc w:val="both"/></w:pPr>';
    expect(pPrOf(pPr, toList(4))).toBe(
      `<w:pPr>${NUM_PR}<w:spacing w:after="240"/>` +
        '<w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/></w:pPr>'
    );
  });

  it("does not touch the indent when told to keep it", () => {
    const pPr = '<w:pPr><w:ind w:left="1440" w:hanging="400"/></w:pPr>';
    expect(
      pPrOf(pPr, { numbering: { numId: 4, ilvl: 0 }, indent: { kind: "keep" } })
    ).toBe(`<w:pPr>${NUM_PR}<w:ind w:left="1440" w:hanging="400"/></w:pPr>`);
  });
});

describe("moving the left indent", () => {
  function indentedPPr(pPr: string | null, leftTwips: number): string | null {
    const props = withLeftIndent(pPr, leftTwips);
    if (!props) throw new Error("the fragment was not recognized");
    return props.pPr;
  }

  it("creates a fresh fragment for a paragraph that had no formatting", () => {
    expect(indentedPPr(null, 720)).toBe('<w:pPr><w:ind w:left="720"/></w:pPr>');
  });

  it("goes into the slot OOXML prescribes (after spacing, before jc)", () => {
    const pPr = '<w:pPr><w:spacing w:after="240"/><w:jc w:val="both"/></w:pPr>';
    expect(indentedPPr(pPr, 720)).toBe(
      '<w:pPr><w:spacing w:after="240"/><w:ind w:left="720"/>' +
        '<w:jc w:val="both"/></w:pPr>'
    );
  });

  it("changes the left indent and leaves the other indents alone", () => {
    const pPr =
      '<w:pPr><w:ind w:left="720" w:right="200" w:hanging="360"/></w:pPr>';
    expect(indentedPPr(pPr, 1440)).toBe(
      '<w:pPr><w:ind w:left="1440" w:right="200" w:hanging="360"/></w:pPr>'
    );
  });

  it("a first-line indent survives the move", () => {
    const pPr = '<w:pPr><w:ind w:left="400" w:firstLine="200"/></w:pPr>';
    expect(indentedPPr(pPr, 1120)).toBe(
      '<w:pPr><w:ind w:left="1120" w:firstLine="200"/></w:pPr>'
    );
  });

  it("keeps the slot the document used for the value", () => {
    const pPr = '<w:pPr><w:ind w:start="720"/></w:pPr>';
    expect(indentedPPr(pPr, 1440)).toBe(
      '<w:pPr><w:ind w:start="1440"/></w:pPr>'
    );
  });

  /** A character-unit indent overrules the twips one in Word, so it cannot be left behind */
  it("drops a left indent measured in characters", () => {
    const pPr = '<w:pPr><w:ind w:leftChars="200" w:right="100"/></w:pPr>';
    expect(indentedPPr(pPr, 720)).toBe(
      '<w:pPr><w:ind w:left="720" w:right="100"/></w:pPr>'
    );
  });

  it("zero writes no left indent, putting the paragraph back where it started", () => {
    const original = '<w:pPr><w:ind w:right="200" w:hanging="360"/></w:pPr>';
    const indented = indentedPPr(original, 720);
    expect(indented).toBe(
      '<w:pPr><w:ind w:left="720" w:right="200" w:hanging="360"/></w:pPr>'
    );
    expect(indentedPPr(indented, 0)).toBe(original);
  });

  it("a fragment left with nothing but an empty indent disappears entirely", () => {
    expect(indentedPPr('<w:pPr><w:ind w:left="720"/></w:pPr>', 0)).toBeNull();
  });

  it("the display values come back out of the operated-on fragment", () => {
    expect(withLeftIndent(null, 1440)?.format).toEqual({ indentStartPt: 72 });
  });
});

describe("setting the line spacing", () => {
  function spacedPPr(pPr: string | null, spacing: LineSpacing): string | null {
    const props = withLineSpacing(pPr, spacing);
    if (!props) throw new Error("the fragment was not recognized");
    return props.pPr;
  }

  const DOUBLE: LineSpacing = { rule: "auto", lines: 2 };

  it("writes the multiple in 240ths of a line", () => {
    expect(spacedPPr(null, { rule: "auto", lines: 1.15 })).toBe(
      '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>'
    );
    expect(spacedPPr(null, DOUBLE)).toBe(
      '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>'
    );
  });

  it("writes a pinned-down height in twips", () => {
    expect(spacedPPr(null, { rule: "exact", pt: 18 })).toBe(
      '<w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr>'
    );
  });

  it("goes into the slot OOXML prescribes (before ind)", () => {
    const pPr = '<w:pPr><w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>';
    expect(spacedPPr(pPr, DOUBLE)).toBe(
      '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/>' +
        '<w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>'
    );
  });

  it("leaves the space above and below the paragraph as it was", () => {
    const pPr =
      '<w:pPr><w:spacing w:before="120" w:after="240" ' +
      'w:line="240" w:lineRule="auto"/></w:pPr>';
    expect(spacedPPr(pPr, DOUBLE)).toBe(
      '<w:pPr><w:spacing w:before="120" w:after="240" ' +
        'w:line="480" w:lineRule="auto"/></w:pPr>'
    );
  });

  it("the display values come back out of the operated-on fragment", () => {
    expect(withLineSpacing(null, DOUBLE)?.format).toEqual({
      lineSpacing: DOUBLE,
    });
  });
});

describe("swapping the alignment in", () => {
  function alignedPPr(
    pPr: string | null,
    align: ParagraphAlign
  ): string | null {
    const props = withParagraphAlign(pPr, align);
    if (!props) throw new Error("the fragment was not recognized");
    return props.pPr;
  }

  it("creates a fresh fragment for a paragraph that had no formatting", () => {
    expect(alignedPPr(null, "center")).toBe(
      '<w:pPr><w:jc w:val="center"/></w:pPr>'
    );
  });

  it("writes justified alignment as both", () => {
    expect(alignedPPr(null, "justify")).toBe(
      '<w:pPr><w:jc w:val="both"/></w:pPr>'
    );
    expect(withParagraphAlign(null, "justify")?.format).toEqual({
      align: "justify",
    });
  });

  it("an existing alignment keeps its slot and only its value changes", () => {
    const pPr =
      '<w:pPr><w:jc w:val="both"/>' +
      '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>';
    expect(alignedPPr(pPr, "right")).toBe(
      '<w:pPr><w:jc w:val="right"/>' +
        '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
    );
  });

  it("an alignment that was not there goes into the slot OOXML prescribes (after ind, before rPr)", () => {
    const pPr =
      '<w:pPr><w:ind w:left="720"/>' + "<w:rPr><w:b/></w:rPr></w:pPr>";
    expect(alignedPPr(pPr, "left")).toBe(
      '<w:pPr><w:ind w:left="720"/><w:jc w:val="left"/>' +
        "<w:rPr><w:b/></w:rPr></w:pPr>"
    );
  });

  it("formatting we do not read and the list position stay as they are", () => {
    const pPr = `<w:pPr>${NUM_PR}<w:spacing w:line="276" w:lineRule="auto"/></w:pPr>`;
    expect(alignedPPr(pPr, "center")).toBe(
      `<w:pPr>${NUM_PR}<w:spacing w:line="276" w:lineRule="auto"/>` +
        '<w:jc w:val="center"/></w:pPr>'
    );
  });

  it("display values that came from a style survive fixing the alignment", () => {
    const styles = readStyles(
      parseXml(
        `<w:styles xmlns:w="${W_NS}"><w:style w:styleId="Item">` +
          '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
          "</w:style></w:styles>"
      )
    );
    const props = withParagraphAlign(
      '<w:pPr><w:pStyle w:val="Item"/></w:pPr>',
      "right",
      styles
    );
    // The alignment the paragraph wrote down beats the style
    expect(props?.format).toEqual({ align: "right", spaceAfterPt: 12 });
  });
});

describe("pointing the paragraph at a style", () => {
  function styledPPr(pPr: string | null, styleId: string | null) {
    const props = withParagraphStyle(pPr, styleId);
    if (!props) throw new Error("the fragment was not recognized");
    return props.pPr;
  }

  it("creates a fresh fragment for a paragraph that had no formatting", () => {
    expect(styledPPr(null, "Heading1")).toBe(
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    );
  });

  it("a style name that was not there leads the fragment, the slot OOXML prescribes", () => {
    const pPr = '<w:pPr><w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>';
    expect(styledPPr(pPr, "Heading1")).toBe(
      '<w:pPr><w:pStyle w:val="Heading1"/>' +
        '<w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>'
    );
  });

  it("an existing style name only has its value changed", () => {
    const pPr = '<w:pPr><w:pStyle w:val="Quote"/><w:jc w:val="both"/></w:pPr>';
    expect(styledPPr(pPr, "Heading1")).toBe(
      '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="both"/></w:pPr>'
    );
  });

  it("a null id takes the style name away and leaves the rest of the formatting", () => {
    const pPr = '<w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="720"/></w:pPr>';
    expect(styledPPr(pPr, null)).toBe('<w:pPr><w:ind w:left="720"/></w:pPr>');
  });

  it("there is no fragment left once the style name was all it held", () => {
    expect(
      withParagraphStyle('<w:pPr><w:pStyle w:val="Quote"/></w:pPr>', null)
    ).toEqual({ pPr: null, format: null, styleRun: null });
  });

  it("the display values are read again under the new style", () => {
    const styles = readStyles(
      parseXml(
        `<w:styles xmlns:w="${W_NS}">` +
          '<w:style w:styleId="Quote"><w:pPr><w:jc w:val="center"/>' +
          '<w:spacing w:after="240"/></w:pPr></w:style>' +
          '<w:style w:styleId="Heading1"><w:pPr><w:spacing w:before="480"/>' +
          "</w:pPr></w:style></w:styles>"
      )
    );
    const pPr = '<w:pPr><w:pStyle w:val="Quote"/><w:jc w:val="right"/></w:pPr>';
    const props = withParagraphStyle(pPr, "Heading1", styles);

    // The values of the old style are gone and the new one's are underneath,
    // with the alignment the paragraph wrote down still on top
    expect(props?.format).toEqual({ align: "right", spaceBeforePt: 24 });
  });

  it("writes a name holding markup characters as text", () => {
    expect(styledPPr(null, 'A & B "C"')).toBe(
      '<w:pPr><w:pStyle w:val="A &amp; B &quot;C&quot;"/></w:pPr>'
    );
  });
});

describe("reading the display values again", () => {
  it("the display values come back out of the operated-on fragment", () => {
    const props = withListNumbering(
      '<w:pPr><w:jc w:val="center"/></w:pPr>',
      toList(4, 1)
    );
    expect(props?.format).toEqual({
      align: "center",
      numbering: { numId: 4, ilvl: 1 },
      indentStartPt: 72,
      textIndentPt: -18,
    });
  });

  it("leaving the list makes the list disappear from the display values too", () => {
    const props = withListNumbering(
      `<w:pPr>${NUM_PR}<w:ind w:left="720" w:hanging="360"/></w:pPr>`,
      LEAVE_LIST
    );
    expect(props?.format).toEqual({ indentStartPt: 36 });
  });

  it("there are no display values either once the fragment is gone", () => {
    const props = withListNumbering(`<w:pPr>${NUM_PR}</w:pPr>`, LEAVE_LIST);
    expect(props).toEqual({ pPr: null, format: null, styleRun: null });
  });
});

describe("a paragraph that points at a style", () => {
  const pPr = '<w:pPr><w:pStyle w:val="Item"/></w:pPr>';

  /** The `Item` style passes down the alignment and the space after the paragraph */
  function styleTable(): StyleTable {
    return readStyles(
      parseXml(
        `<w:styles xmlns:w="${W_NS}"><w:style w:styleId="Item">` +
          '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
          "</w:style></w:styles>"
      )
    );
  }

  it("display values that came from a style survive fixing the list", () => {
    const props = withListNumbering(pPr, toList(4), styleTable());
    expect(props?.format).toEqual({
      align: "center",
      spaceAfterPt: 12,
      numbering: { numId: 4, ilvl: 0 },
      indentStartPt: 36,
      textIndentPt: -18,
    });
  });

  it("with no style table it reads only what the paragraph wrote down", () => {
    const props = withListNumbering(pPr, toList(4));
    expect(props?.format).toEqual({
      numbering: { numId: 4, ilvl: 0 },
      indentStartPt: 36,
      textIndentPt: -18,
    });
  });

  it("the styleId it points at stays in the fragment as it is", () => {
    expect(pPrOf(pPr, toList(4))).toBe(
      '<w:pPr><w:pStyle w:val="Item"/>' +
        `${NUM_PR}<w:ind w:left="720" w:hanging="360"/></w:pPr>`
    );
  });
});

describe("a fragment none of the operations recognize", () => {
  const DOUBLE: LineSpacing = { rule: "auto", lines: 2 };

  it.each([
    [
      "putting the paragraph in a list",
      (pPr: string) => withListNumbering(pPr, toList(4)),
    ],
    ["moving the left indent", (pPr: string) => withLeftIndent(pPr, 720)],
    ["setting the line spacing", (pPr: string) => withLineSpacing(pPr, DOUBLE)],
    [
      "swapping the alignment in",
      (pPr: string) => withParagraphAlign(pPr, "center"),
    ],
    [
      "pointing the paragraph at a style",
      (pPr: string) => withParagraphStyle(pPr, "Quote"),
    ],
  ])("%s leaves it untouched", (_name, operate) => {
    expect(operate("<w:pPr>unclosed")).toBeNull();
  });
});
