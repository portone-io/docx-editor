// @vitest-environment jsdom
/**
 * Which cells a border or fill job acts on, and which of their sides it writes.
 *
 * The XML surgery itself is pinned in `docx/tableFormatting.test`; what is checked here is that a
 * job reaches exactly the selected cells and leaves every other cell holding its original XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { runCommand } from "../__testing__/editing";
import { toCellFormat } from "../model/format";
import { cellStyle } from "../styles/inlineStyle";
import {
  cell,
  cellWithText,
  dxa,
  firstTable,
  row,
  schema,
  stateWithCellsSelected,
  stateWithCursorIn,
  tableDoc,
} from "./__testing__/tables";
import {
  activeCellBackground,
  activeCellBorderColor,
  activeCellPadding,
  activeCellVerticalAlign,
  canSetCellBorderColor,
  setCellBackground,
  setCellBorderColor,
  setCellBorders,
  setCellPadding,
  setCellVerticalAlign,
} from "./cellFormatting";

/** A table that draws the lines between its cells from the table level, as a contract table does */
const INSIDE_LINES =
  "<w:tblPr><w:tblBorders>" +
  '<w:insideH w:val="single" w:sz="4" w:color="999999"/>' +
  '<w:insideV w:val="single" w:sz="4" w:color="999999"/>' +
  "</w:tblBorders></w:tblPr>";

/**
 * The display values the import gives every cell of such a table: the table's inside lines are
 * folded in as each cell's four sides, even though no cell wrote a border of its own
 */
const INSIDE_LINE_FORMAT = {
  borderTop: "0.5pt solid #999999",
  borderBottom: "0.5pt solid #999999",
  borderLeft: "0.5pt solid #999999",
  borderRight: "0.5pt solid #999999",
};

/** A 3x3 table whose cells are named by column and row (`b2` is the middle one) */
function gridDoc(tblPr: string | null = null): PMNode {
  const format = tblPr === null ? null : INSIDE_LINE_FORMAT;
  const line = (...texts: string[]) =>
    row(...texts.map((text) => cell(text, { tcW: dxa(3000), format })));
  return tableDoc(
    [line("a1", "b1", "c1"), line("a2", "b2", "c2"), line("a3", "b3", "c3")],
    tblPr === null ? null : { tblPr }
  );
}

function tcPrOf(doc: PMNode, text: string): unknown {
  return cellWithText(firstTable(doc).table, text)?.attrs.tcPr;
}

function formatOf(doc: PMNode, text: string) {
  return toCellFormat(cellWithText(firstTable(doc).table, text)?.attrs.format);
}

describe("the cells a formatting job acts on", () => {
  it("is the cell holding the caret, and no other", () => {
    const edited = runCommand(
      stateWithCursorIn(gridDoc(), "b2"),
      setCellBorders("all")
    );
    expect(tcPrOf(edited.doc, "b2")).toContain("<w:tcBorders>");
    for (const other of ["a1", "b1", "c1", "a2", "c2", "a3", "b3", "c3"]) {
      expect(tcPrOf(edited.doc, other)).toBeNull();
    }
  });

  it("is every cell of a cell selection", () => {
    const edited = runCommand(
      stateWithCellsSelected(gridDoc(), "a1", "b2"),
      setCellBorders("all")
    );
    for (const inside of ["a1", "b1", "a2", "b2"]) {
      expect(tcPrOf(edited.doc, inside)).toContain("<w:tcBorders>");
    }
    for (const outside of ["c1", "c2", "a3", "b3", "c3"]) {
      expect(tcPrOf(edited.doc, outside)).toBeNull();
    }
  });

  it("is nothing at all outside a table", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("outside")),
    ]);
    const state = EditorState.create({ doc, schema });
    expect(setCellBorders("all")(state)).toBe(false);
    expect(setCellBackground("#ffff00")(state)).toBe(false);
    expect(setCellBorderColor("#ff0000")(state)).toBe(false);
    expect(canSetCellBorderColor(state)).toBe(false);
    expect(activeCellBackground(state)).toBeNull();
  });
});

describe("cell alignment and padding", () => {
  it("writes vertical alignment to every selected cell and reports mixed values", () => {
    const selected = stateWithCellsSelected(gridDoc(), "a1", "b1");
    const centered = runCommand(selected, setCellVerticalAlign("center"));
    expect(activeCellVerticalAlign(centered)).toBe("center");
    expect(tcPrOf(centered.doc, "a1")).toContain('<w:vAlign w:val="center"/>');

    const mixed = stateWithCellsSelected(
      runCommand(
        stateWithCursorIn(gridDoc(), "a1"),
        setCellVerticalAlign("bottom")
      ).doc,
      "a1",
      "b1"
    );
    expect(activeCellVerticalAlign(mixed)).toBe("mixed");
  });

  it("writes only supplied padding sides and reports mixed sides independently", () => {
    const leftOnly = runCommand(
      stateWithCursorIn(gridDoc(), "a1"),
      setCellPadding({ left: 12 })
    );
    expect(tcPrOf(leftOnly.doc, "a1")).toContain(
      '<w:tcMar><w:left w:w="240" w:type="dxa"/></w:tcMar>'
    );
    expect(formatOf(leftOnly.doc, "a1")?.paddingLeftPt).toBe(12);

    const selected = stateWithCellsSelected(leftOnly.doc, "a1", "b1");
    expect(activeCellPadding(selected)).toEqual({
      top: null,
      right: null,
      bottom: null,
      left: "mixed",
    });
  });

  it("updates strict leading margins through the left-to-right API", () => {
    const doc = tableDoc([
      row(
        cell("strict", {
          tcPr:
            '<w:tcPr><w:tcMar><w:start w:w="80" w:type="dxa"/>' +
            '<w:end w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>',
        })
      ),
    ]);
    const edited = runCommand(
      stateWithCursorIn(doc, "strict"),
      setCellPadding({ left: 8 })
    );

    expect(tcPrOf(edited.doc, "strict")).toContain(
      '<w:start w:w="160" w:type="dxa"/>'
    );
    expect(tcPrOf(edited.doc, "strict")).not.toContain("<w:left");
    expect(formatOf(edited.doc, "strict")?.paddingLeftPt).toBe(8);
    expect(activeCellPadding(edited).left).toBe(8);
  });

  it("reports and performs no layout formatting inside a locked cell", () => {
    const doc = tableDoc([
      row(cell("locked", { sdtContentsLocked: true, sdtDeletionLocked: true })),
    ]);
    const state = stateWithCursorIn(doc, "locked");
    expect(setCellVerticalAlign("center")(state)).toBe(false);
    expect(setCellPadding({ top: 6 })(state)).toBe(false);
  });
});

describe("the border presets", () => {
  it("draw every side of every selected cell", () => {
    const edited = runCommand(
      stateWithCursorIn(gridDoc(), "b2"),
      setCellBorders("all")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      borderTop: "0.5pt solid #000000",
      borderBottom: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      borderRight: "0.5pt solid #000000",
    });
  });

  it("draw only the edge of the selected block, leaving the cells inside it alone", () => {
    const edited = runCommand(
      stateWithCellsSelected(gridDoc(), "a1", "c3"),
      setCellBorders("outer")
    );

    // The corner takes two sides, the cell on one edge takes one, and the middle cell is untouched
    expect(tcPrOf(edited.doc, "a1")).toContain("<w:top ");
    expect(tcPrOf(edited.doc, "a1")).toContain("<w:left ");
    expect(tcPrOf(edited.doc, "a1")).not.toContain("<w:right ");
    expect(tcPrOf(edited.doc, "b1")).toContain("<w:top ");
    expect(tcPrOf(edited.doc, "b1")).not.toContain("<w:bottom ");
    expect(tcPrOf(edited.doc, "b2")).toBeNull();
  });

  it("switch a cell's own lines off where the table draws them", () => {
    const edited = runCommand(
      stateWithCursorIn(gridDoc(INSIDE_LINES), "b2"),
      setCellBorders("none")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      borderTop: "none",
      borderBottom: "none",
      borderLeft: "none",
      borderRight: "none",
    });
    // The neighbours keep the lines the table lays down
    expect(formatOf(edited.doc, "b1")?.borderBottom).toBe(
      "0.5pt solid #999999"
    );
  });
});

describe("the border color", () => {
  it("colors visible lines inherited from the table with direct cell overrides", () => {
    const state = stateWithCursorIn(gridDoc(INSIDE_LINES), "b2");
    expect(canSetCellBorderColor(state)).toBe(true);
    const colored = runCommand(state, setCellBorderColor("#ff0000"));

    expect(tcPrOf(colored.doc, "b2")).toContain(
      '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/>'
    );
    expect(formatOf(colored.doc, "b2")).toMatchObject({
      borderTop: "0.5pt solid #FF0000",
      borderBottom: "0.5pt solid #FF0000",
      borderLeft: "0.5pt solid #FF0000",
      borderRight: "0.5pt solid #FF0000",
    });
    expect(formatOf(colored.doc, "b1")?.borderBottom).toBe(
      "0.5pt solid #999999"
    );
    expect(firstTable(colored.doc).table.attrs.tblPr).toBe(INSIDE_LINES);
  });

  it("reports unavailable when any selected cell is locked", () => {
    const ownBorder =
      '<w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4"/>' +
      "</w:tcBorders></w:tcPr>";
    const doc = tableDoc([
      row(
        cell("locked", {
          tcPr: ownBorder,
          sdtContentsLocked: true,
          sdtDeletionLocked: true,
        }),
        cell("plain", { tcPr: ownBorder })
      ),
    ]);
    const state = stateWithCellsSelected(doc, "locked", "plain");

    expect(canSetCellBorderColor(state)).toBe(false);
    expect(setCellBorderColor("#ff0000")(state)).toBe(false);
  });

  it("colors the lines of every selected cell", () => {
    const drawn = runCommand(
      stateWithCellsSelected(gridDoc(), "a1", "b1"),
      setCellBorders("all")
    );
    const colored = runCommand(drawn, setCellBorderColor("#ff0000"));

    expect(formatOf(colored.doc, "a1")?.borderTop).toBe("0.5pt solid #FF0000");
    expect(formatOf(colored.doc, "b1")?.borderRight).toBe(
      "0.5pt solid #FF0000"
    );
    expect(activeCellBorderColor(colored)).toBe("#FF0000");
  });

  it("reports the color only when every visible line shares it", () => {
    const lines = stateWithCursorIn(gridDoc(INSIDE_LINES), "b2");
    expect(activeCellBorderColor(lines)).toBe("#999999");

    const drawn = runCommand(lines, setCellBorders("all"));
    // Now black on the cell's own sides against the table's grey inside lines nowhere to be seen
    expect(activeCellBorderColor(drawn)).toBe("#000000");
  });
});

describe("the lines a table style laid down", () => {
  /**
   * The lines of a table that draws none itself: they come from the table style it points at, so
   * the tblPr says nothing about them and the table carries them as `styleInside` instead
   */
  const STYLE_INSIDE = {
    horizontal: "0.5pt solid #999999",
    vertical: "0.5pt solid #999999",
  };

  /** The same 3x3 grid, with the lines coming from the style rather than from the tblPr */
  function styledGridDoc(): PMNode {
    const line = (...texts: string[]) =>
      row(
        ...texts.map((text) =>
          cell(text, { tcW: dxa(3000), format: INSIDE_LINE_FORMAT })
        )
      );
    return tableDoc(
      [line("a1", "b1", "c1"), line("a2", "b2", "c2"), line("a3", "b3", "c3")],
      { styleInside: STYLE_INSIDE }
    );
  }

  it("survives a fill written on the cell", () => {
    const edited = runCommand(
      stateWithCursorIn(styledGridDoc(), "b2"),
      setCellBackground("#ffff00")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      background: "#FFFF00",
      ...INSIDE_LINE_FORMAT,
    });
  });

  it("stays on the sides an outer preset does not write", () => {
    const edited = runCommand(
      stateWithCellsSelected(styledGridDoc(), "a1", "c3"),
      setCellBorders("outer")
    );
    expect(formatOf(edited.doc, "a1")).toEqual({
      borderTop: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      // The two sides facing the inside of the block keep the line the style laid down
      borderBottom: "0.5pt solid #999999",
      borderRight: "0.5pt solid #999999",
    });
  });

  it("is switched off where the cell itself says to draw nothing", () => {
    const edited = runCommand(
      stateWithCursorIn(styledGridDoc(), "b2"),
      setCellBorders("none")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      borderTop: "none",
      borderBottom: "none",
      borderLeft: "none",
      borderRight: "none",
    });
  });
});

describe("the cell padding the table lays down", () => {
  /** The margins a table style laid down, as the import records them on the table node */
  const STYLE_MARGINS = {
    topPt: 0,
    rightPt: 5.4,
    bottomPt: 0,
    leftPt: 5.4,
  };

  /** The padding those margins give every cell */
  const PADDING = {
    paddingTopPt: 0,
    paddingRightPt: 5.4,
    paddingBottomPt: 0,
    paddingLeftPt: 5.4,
  };

  /** The same 3x3 grid, with the padding coming from wherever the table attrs put it */
  function paddedGridDoc(tableAttrs: Record<string, unknown>): PMNode {
    const line = (...texts: string[]) =>
      row(
        ...texts.map((text) => cell(text, { tcW: dxa(3000), format: PADDING }))
      );
    return tableDoc(
      [line("a1", "b1", "c1"), line("a2", "b2", "c2"), line("a3", "b3", "c3")],
      tableAttrs
    );
  }

  const BLACK_LINES = {
    borderTop: "0.5pt solid #000000",
    borderBottom: "0.5pt solid #000000",
    borderLeft: "0.5pt solid #000000",
    borderRight: "0.5pt solid #000000",
  };

  it("survives a border drawn on the cell, when it came from the style", () => {
    const edited = runCommand(
      stateWithCursorIn(
        paddedGridDoc({ styleCellMargins: STYLE_MARGINS }),
        "b2"
      ),
      setCellBorders("all")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      ...BLACK_LINES,
      ...PADDING,
    });
  });

  it("survives it just the same when it came from the tblPr", () => {
    const tblPr =
      "<w:tblPr><w:tblCellMar>" +
      '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
      '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>' +
      "</w:tblCellMar></w:tblPr>";
    const edited = runCommand(
      stateWithCellsSelected(paddedGridDoc({ tblPr }), "a1", "b1"),
      setCellBackground("#ffff00")
    );
    expect(formatOf(edited.doc, "a1")).toEqual({
      background: "#FFFF00",
      ...PADDING,
    });
  });

  it("gives way on the sides the cell wrote a margin of its own", () => {
    const withOwnMargin = {
      ...PADDING,
      paddingLeftPt: 20,
    };
    const doc = tableDoc(
      [
        row(
          cell("a1", {
            tcPr: '<w:tcPr><w:tcMar><w:left w:w="400" w:type="dxa"/></w:tcMar></w:tcPr>',
            format: withOwnMargin,
          }),
          cell("b1", { format: PADDING })
        ),
      ],
      { styleCellMargins: STYLE_MARGINS }
    );
    const edited = runCommand(
      stateWithCursorIn(doc, "a1"),
      setCellBorders("all")
    );
    expect(formatOf(edited.doc, "a1")).toEqual({
      ...BLACK_LINES,
      ...withOwnMargin,
    });
  });
});

describe("the border around the table", () => {
  /** A table drawing all six of its sides: a black line around the outside, grey ones between cells */
  const ALL_LINES =
    "<w:tblPr><w:tblBorders>" +
    ["top", "left", "bottom", "right"]
      .map((side) => `<w:${side} w:val="single" w:sz="4" w:color="000000"/>`)
      .join("") +
    '<w:insideH w:val="single" w:sz="4" w:color="999999"/>' +
    '<w:insideV w:val="single" w:sz="4" w:color="999999"/>' +
    "</w:tblBorders></w:tblPr>";

  /** The outer lines as the import records them on the table node */
  const OUTER = {
    borderTop: "0.5pt solid #000000",
    borderBottom: "0.5pt solid #000000",
    borderLeft: "0.5pt solid #000000",
    borderRight: "0.5pt solid #000000",
  };

  /** The same 3x3 grid, with a line around the outside as well as between the cells */
  function borderedGridDoc(): PMNode {
    const line = (...texts: string[]) =>
      // A job derives the display values of the cells it touches again, so they start out empty
      row(...texts.map((text) => cell(text, { tcW: dxa(3000) })));
    return tableDoc(
      [line("a1", "b1", "c1"), line("a2", "b2", "c2"), line("a3", "b3", "c3")],
      { tblPr: ALL_LINES, format: OUTER }
    );
  }

  it("stays under a cell on the edge of the grid when it is filled", () => {
    const edited = runCommand(
      stateWithCursorIn(borderedGridDoc(), "a1"),
      setCellBackground("#ffff00")
    );
    expect(formatOf(edited.doc, "a1")).toEqual({
      background: "#FFFF00",
      borderTop: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      // The two sides facing the rest of the grid draw the line between the cells
      borderBottom: "0.5pt solid #999999",
      borderRight: "0.5pt solid #999999",
    });
  });

  it("is switched off by the no-border preset, the same as the lines inside", () => {
    const edited = runCommand(
      stateWithCellsSelected(borderedGridDoc(), "a1", "c3"),
      setCellBorders("none")
    );
    for (const text of ["a1", "b2", "c3"]) {
      expect(formatOf(edited.doc, text)).toEqual({
        borderTop: "none",
        borderBottom: "none",
        borderLeft: "none",
        borderRight: "none",
      });
    }
  });
});

describe("the cell fill", () => {
  it("is written on every selected cell and drawn from the cell's display values", () => {
    const edited = runCommand(
      stateWithCellsSelected(gridDoc(), "a1", "a2"),
      setCellBackground("#ffff00")
    );
    expect(tcPrOf(edited.doc, "a1")).toContain('w:fill="FFFF00"');
    expect(cellStyle(formatOf(edited.doc, "a2"))).toBe(
      "background-color:#FFFF00"
    );
    expect(activeCellBackground(edited)).toBe("#FFFF00");
  });

  it("leaves the lines the table draws in place", () => {
    const edited = runCommand(
      stateWithCursorIn(gridDoc(INSIDE_LINES), "b2"),
      setCellBackground("#ffff00")
    );
    expect(formatOf(edited.doc, "b2")).toEqual({
      background: "#FFFF00",
      borderTop: "0.5pt solid #999999",
      borderBottom: "0.5pt solid #999999",
      borderLeft: "0.5pt solid #999999",
      borderRight: "0.5pt solid #999999",
    });
  });

  it("is taken back off by the no-fill entry", () => {
    const filled = runCommand(
      stateWithCursorIn(gridDoc(), "b2"),
      setCellBackground("#ffff00")
    );
    const cleared = runCommand(filled, setCellBackground(null));

    expect(tcPrOf(cleared.doc, "b2")).toBeNull();
    expect(formatOf(cleared.doc, "b2")).toBeNull();
    expect(activeCellBackground(cleared)).toBeNull();
  });

  it("reports no shared fill when the selected cells differ", () => {
    const filled = runCommand(
      stateWithCursorIn(gridDoc(), "a1"),
      setCellBackground("#ffff00")
    );
    const selected = stateWithCellsSelected(filled.doc, "a1", "b1");
    expect(activeCellBackground(selected)).toBeNull();
  });
});
