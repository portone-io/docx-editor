// @vitest-environment node
import { describe, expect, it } from "vitest";
import { DEFAULT_FONT_FALLBACKS } from "./fontStack";
import {
  cellStyle,
  documentDefaultsStyle,
  lineHeightValue,
  paragraphStyle,
  rowStyle,
  runStyle,
  SINGLE_LINE_RATIO,
  tableStyle,
} from "./inlineStyle";

const DEFAULT_FONT_STACK = DEFAULT_FONT_FALLBACKS.defaultStack;

describe("font fallback", () => {
  it("appends the same family after a run's font", () => {
    const css = runStyle({ fontFamily: '"Malgun Gothic"' });
    expect(css).toContain('font-family:"Malgun Gothic",');
    expect(css).toContain("Pretendard");
  });

  it("the document default font goes through the same path", () => {
    const css = documentDefaultsStyle({
      fontSizePt: null,
      fontFamily: '"Batang"',
      lineSpacing: null,
    });
    expect(css).toContain('--docx-editor-font-family:"Batang",');
    expect(css).toContain('"Noto Serif KR"');
  });

  it("settles on a font even when the document does not write down a default one", () => {
    const css = documentDefaultsStyle({
      fontSizePt: null,
      fontFamily: null,
      lineSpacing: null,
    });
    expect(css).toContain(`--docx-editor-font-family:${DEFAULT_FONT_STACK}`);
  });
});

describe("the lines of a table", () => {
  const BLACK = "0.5pt solid #000000";
  const AROUND = {
    borderTop: BLACK,
    borderBottom: BLACK,
    borderLeft: BLACK,
    borderRight: BLACK,
  };

  /**
   * Every line of a table is drawn by its cells. A border on the table element would win over a
   * cell that says to draw nothing, because CSS resolves a collapsed border in favour of the line.
   */
  it("are left to the cells, so the table itself draws none of them", () => {
    expect(tableStyle({ ...AROUND, background: "#FFFF00" }, null)).toBe(
      "background-color:#FFFF00"
    );
    expect(tableStyle(AROUND, { type: "dxa", twips: 4500 })).toBe(
      "width:300px"
    );
  });

  it("are drawn by the cell, one side at a time", () => {
    expect(cellStyle({ ...AROUND, borderTop: "none" })).toBe(
      `border-top:none;border-bottom:${BLACK};` +
        `border-left:${BLACK};border-right:${BLACK}`
    );
  });
});

describe("a width that leaves the size to the layout", () => {
  it("asks for no width at all, so the layout decides", () => {
    expect(tableStyle(null, { type: "auto" })).toBeUndefined();
    expect(tableStyle(null, { type: "nil" })).toBeUndefined();
    expect(tableStyle(null, { type: "pct", fiftieths: 2500 })).toBe(
      "width:50%"
    );
  });
});

describe("a fill that paints nothing", () => {
  /**
   * A document that took a fill off draws transparent rather than nothing at all, so that it wins
   * over the fill lying behind it
   */
  it("is drawn as transparent, wherever it appears", () => {
    expect(cellStyle({ background: "none" })).toBe(
      "background-color:transparent"
    );
    expect(tableStyle({ background: "none" }, null)).toBe(
      "background-color:transparent"
    );
    expect(paragraphStyle({ background: "none" })).toBe(
      "background-color:transparent"
    );
    expect(runStyle({ background: "none" })).toBe(
      "background-color:transparent"
    );
  });
});

describe("paragraph direction", () => {
  it("uses the logical direction recorded by the paragraph", () => {
    expect(paragraphStyle({ direction: "rtl" })).toBe("direction:rtl");
  });

  it("keeps Strict paragraph indents on their logical sides", () => {
    expect(
      paragraphStyle({
        direction: "rtl",
        indentStartPt: 30,
        indentEndPt: 15,
      })
    ).toBe("direction:rtl;margin-inline-start:30pt;margin-inline-end:15pt");
  });
});

describe("how a table is placed on the page", () => {
  it("keeps its alignment when it stands at the margin, indent written or not", () => {
    const centered = "margin-left:auto;margin-right:auto";
    expect(tableStyle({ align: "center" }, null)).toBe(centered);
    // Almost every table writes an indent of 0, which is not a reason to stop centering it
    expect(tableStyle({ align: "center", indentLeftPt: 0 }, null)).toBe(
      centered
    );
  });

  it("draws the indent instead once it pushes the table off the margin", () => {
    expect(tableStyle({ align: "center", indentLeftPt: 36 }, null)).toBe(
      "margin-left:36pt"
    );
    expect(tableStyle({ indentLeftPt: -5.4 }, null)).toBe("margin-left:-5.4pt");
  });
});

describe("the height of a row", () => {
  it("is drawn the same whichever rule the document wrote", () => {
    expect(rowStyle({ height: { rule: "atLeast", pt: 27 } })).toBe(
      "height:27pt"
    );
    expect(rowStyle({ height: { rule: "exact", pt: 27 } })).toBe("height:27pt");
    expect(rowStyle({})).toBeUndefined();
    expect(rowStyle(null)).toBeUndefined();
  });
});

describe("line spacing", () => {
  it("one line's height is 1.3 times the font size", () => {
    expect(SINGLE_LINE_RATIO).toBe(1.3);
    expect(lineHeightValue(null)).toBe("1.3");
  });

  it("auto multiplies the line count by one line's height", () => {
    // Measured from the fixtures. w:line="276" is 1.15 lines and "270" is 1.125 lines
    expect(lineHeightValue({ rule: "auto", lines: 1.15 })).toBe("1.5");
    expect(lineHeightValue({ rule: "auto", lines: 1.125 })).toBe("1.46");
  });

  it("uses a pinned height as is, and measures atLeast against one line's height", () => {
    expect(lineHeightValue({ rule: "exact", pt: 14 })).toBe("14pt");
    expect(lineHeightValue({ rule: "atLeast", pt: 14 })).toBe(
      "max(14pt,1.3em)"
    );
  });
});
