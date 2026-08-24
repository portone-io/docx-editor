// @vitest-environment jsdom
/**
 * The lines of a table's cells once an edit has moved its grid around.
 *
 * The tables are built from real XML and mounted in a full editing state, so the whole chain is
 * checked: the formatting the import derived, the structure edit, and the plugin behind it.
 * They draw 1pt around the outside and 0.5pt between the cells, so a line left in the wrong place
 * shows up: after every edit, only the sides on the edge of the grid may draw the outer line.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { runCommand } from "../__testing__/editing";
import { buildTable } from "../docx/importTable";
import { createEditorState } from "../editor/createEditor";
import { type CellFormat, toCellFormat } from "../model/format";
import { parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { setCellBorderColor } from "./cellFormatting";
import {
  addColumnAfter,
  addRowAfter,
  deleteColumn,
  deleteRow,
  mergeCells,
  splitCell,
  type TableCommand,
} from "./commands";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** `w:sz` counts eighths of a point, so 8 is 1pt and 4 is 0.5pt */
const borderXml = (side: string, eighths: number) =>
  `<w:${side} w:val="single" w:sz="${eighths}" w:color="000000"/>`;

/** A table with a 1pt line around the outside and 0.5pt lines between its cells */
const TBL_PR =
  "<w:tblPr><w:tblBorders>" +
  ["top", "left", "bottom", "right"]
    .map((side) => borderXml(side, 8))
    .join("") +
  borderXml("insideH", 4) +
  borderXml("insideV", 4) +
  "</w:tblBorders></w:tblPr>";

const OUTER = "1pt solid #000000";
const INNER = "0.5pt solid #000000";

const COLUMN_DXA = 1000;

const grid = (cols: number) =>
  `<w:tblGrid>${`<w:gridCol w:w="${COLUMN_DXA}"/>`.repeat(cols)}</w:tblGrid>`;

interface CellXml {
  /** How many grid columns the cell covers */
  span?: number;
  /** Formatting of the cell's own, written inside its `w:tcPr` after the width */
  props?: string;
}

const cellXml = (text: string, { span = 1, props = "" }: CellXml = {}) =>
  "<w:tc><w:tcPr>" +
  `<w:tcW w:w="${COLUMN_DXA * span}" w:type="dxa"/>` +
  (span > 1 ? `<w:gridSpan w:val="${span}"/>` : "") +
  props +
  "</w:tcPr>" +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const rowXml = (...cells: string[]) => `<w:tr>${cells.join("")}</w:tr>`;

/** A document holding a single table of `cols` grid columns, one row per fragment */
function tableDoc(cols: number, ...rows: string[]): PMNode {
  const xml = `<w:tbl>${TBL_PR}${grid(cols)}${rows.join("")}</w:tbl>`;
  const wrapped = parseXml(`<w:wrap ${W_NS}>${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  const table = el ? buildTable(el, 0) : null;
  if (!table) throw new Error("the table could not be modelled");
  return docxSchema.nodes.doc.create(null, [table]);
}

/** A 3x3 table whose cells are named by column and row (`b2` is the middle one) */
function gridDoc(): PMNode {
  const line = (...texts: string[]) => rowXml(...texts.map((t) => cellXml(t)));
  return tableDoc(
    3,
    line("a1", "b1", "c1"),
    line("a2", "b2", "c2"),
    line("a3", "b3", "c3")
  );
}

function firstTable(doc: PMNode): PMNode {
  const table = doc.child(0);
  if (table.type.name !== "table") throw new Error("no table");
  return table;
}

/** A state with the cursor placed at one grid coordinate of the table */
function stateInCell(doc: PMNode, row: number, col: number): EditorState {
  const table = firstTable(doc);
  const map = TableMap.get(table);
  const state = createEditorState(doc);
  return state.apply(
    state.tr.setSelection(
      TextSelection.near(
        state.doc.resolve(1 + map.map[row * map.width + col] + 1)
      )
    )
  );
}

/** A state with the block between two grid coordinates selected, cell by cell */
function blockSelected(
  doc: PMNode,
  from: [row: number, col: number],
  to: [row: number, col: number]
): EditorState {
  const map = TableMap.get(firstTable(doc));
  const at = ([row, col]: [number, number]) =>
    1 + map.map[row * map.width + col];
  const state = createEditorState(doc);
  return state.apply(
    state.tr.setSelection(CellSelection.create(doc, at(from), at(to)))
  );
}

/** The document with one edit made to the first table, starting from the given cell */
function edit(
  doc: PMNode,
  at: [row: number, col: number],
  command: TableCommand
): PMNode {
  return runCommand(stateInCell(doc, at[0], at[1]), command).doc;
}

function formatAt(table: PMNode, row: number, col: number): CellFormat | null {
  const map = TableMap.get(table);
  return toCellFormat(
    table.nodeAt(map.map[row * map.width + col])?.attrs.format
  );
}

/** The four lines every cell of the table draws, each labelled by the spot the cell starts at */
function linesOf(table: PMNode): string[] {
  const map = TableMap.get(table);
  return Array.from(new Set(map.map)).map((pos) => {
    const format = toCellFormat(table.nodeAt(pos)?.attrs.format);
    const cell = map.findCell(pos);
    return [
      `(${cell.top},${cell.left})`,
      format?.borderTop,
      format?.borderBottom,
      format?.borderLeft,
      format?.borderRight,
    ].join(" ");
  });
}

/**
 * The same, worked out from nothing but where each cell sits: the table's outer line on a side
 * lying on the edge of the grid, and the line between cells on a side facing another cell.
 */
function fromGrid(table: PMNode): string[] {
  const map = TableMap.get(table);
  return Array.from(new Set(map.map)).map((pos) => {
    const cell = map.findCell(pos);
    return [
      `(${cell.top},${cell.left})`,
      cell.top === 0 ? OUTER : INNER,
      cell.bottom === map.height ? OUTER : INNER,
      cell.left === 0 ? OUTER : INNER,
      cell.right === map.width ? OUTER : INNER,
    ].join(" ");
  });
}

describe("the lines of a table the import built", () => {
  it("draw the outer line only where a cell touches the edge of the grid", () => {
    const table = firstTable(gridDoc());
    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 0, 0)).toEqual({
      borderTop: OUTER,
      borderBottom: INNER,
      borderLeft: OUTER,
      borderRight: INNER,
    });
  });
});

describe("adding a row", () => {
  it("hands the bottom line down to the row added under the last one", () => {
    const table = firstTable(edit(gridDoc(), [2, 0], addRowAfter));

    expect(table.childCount).toBe(4);
    expect(linesOf(table)).toEqual(fromGrid(table));
    // The seam between the two last rows is a line between cells, not the table's outer one
    expect(formatAt(table, 2, 0)?.borderBottom).toBe(INNER);
    expect(formatAt(table, 3, 0)?.borderBottom).toBe(OUTER);
  });

  it("leaves the row inserted in the middle of the table with lines between cells", () => {
    const table = firstTable(edit(gridDoc(), [0, 0], addRowAfter));

    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 1, 0)?.borderTop).toBe(INNER);
    expect(formatAt(table, 1, 0)?.borderBottom).toBe(INNER);
    expect(formatAt(table, 1, 0)?.borderLeft).toBe(OUTER);
  });
});

describe("deleting a row", () => {
  it("brings the bottom line back up to the row that is now the last", () => {
    const table = firstTable(edit(gridDoc(), [2, 0], deleteRow));

    expect(table.childCount).toBe(2);
    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 1, 0)?.borderBottom).toBe(OUTER);
  });
});

describe("adding a column", () => {
  it("hands the right hand line over to the column added at the edge", () => {
    const table = firstTable(edit(gridDoc(), [0, 2], addColumnAfter));

    expect(TableMap.get(table).width).toBe(4);
    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 0, 2)?.borderRight).toBe(INNER);
    expect(formatAt(table, 0, 3)?.borderRight).toBe(OUTER);
  });
});

describe("deleting a column", () => {
  it("brings the right hand line back to the column that is now the last", () => {
    const table = firstTable(edit(gridDoc(), [0, 2], deleteColumn));

    expect(TableMap.get(table).width).toBe(2);
    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 0, 1)?.borderRight).toBe(OUTER);
  });
});

describe("merging cells", () => {
  it("gives the merged cell the outer line on the sides it now reaches", () => {
    const doc = gridDoc();
    const table = firstTable(
      runCommand(blockSelected(doc, [0, 0], [0, 2]), mergeCells).doc
    );

    expect(table.child(0).childCount).toBe(1);
    expect(linesOf(table)).toEqual(fromGrid(table));
    // The one cell now covers the whole width, so its right hand side is the table's own
    expect(formatAt(table, 0, 2)?.borderRight).toBe(OUTER);
    expect(formatAt(table, 0, 0)?.borderLeft).toBe(OUTER);
  });
});

describe("splitting a cell", () => {
  /** A table whose first row is a single cell covering the whole width */
  const spannedDoc = () =>
    tableDoc(
      3,
      rowXml(cellXml("Head", { span: 3 })),
      rowXml(cellXml("a2"), cellXml("b2"), cellXml("c2"))
    );

  it("shares the outer sides out again among the cells the split produced", () => {
    const before = firstTable(spannedDoc());
    // The one cell covers the whole width, so both of its sides are on the edge of the grid
    expect(formatAt(before, 0, 0)?.borderLeft).toBe(OUTER);
    expect(formatAt(before, 0, 0)?.borderRight).toBe(OUTER);

    const table = firstTable(edit(spannedDoc(), [0, 0], splitCell));

    expect(table.child(0).childCount).toBe(3);
    expect(linesOf(table)).toEqual(fromGrid(table));
    expect(formatAt(table, 0, 0)?.borderRight).toBe(INNER);
    expect(formatAt(table, 0, 1)?.borderLeft).toBe(INNER);
    expect(formatAt(table, 0, 1)?.borderRight).toBe(INNER);
    expect(formatAt(table, 0, 2)?.borderRight).toBe(OUTER);
  });
});

describe("what a cell wrote down itself", () => {
  /** A 2x2 table whose top left cell draws a 3pt red line of its own on two sides */
  const ownBorderDoc = () => {
    const own =
      "<w:tcBorders>" +
      '<w:top w:val="single" w:sz="24" w:color="FF0000"/>' +
      '<w:bottom w:val="single" w:sz="24" w:color="FF0000"/>' +
      "</w:tcBorders>";
    return tableDoc(
      2,
      rowXml(cellXml("a1", { props: own }), cellXml("b1")),
      rowXml(cellXml("a2"), cellXml("b2"))
    );
  };

  const RED = "3pt solid #FF0000";

  it("resolves inherited neighbours before the initial editor view is drawn", () => {
    const table = firstTable(createEditorState(ownBorderDoc()).doc);
    const map = TableMap.get(table);
    const below = table.nodeAt(map.map[map.width]);

    expect(formatAt(table, 1, 0)?.borderTop).toBe("none");
    expect(below?.attrs.tcPr).not.toContain("FF0000");
  });

  it("survives a row being added under it", () => {
    const table = firstTable(edit(ownBorderDoc(), [1, 0], addRowAfter));

    expect(formatAt(table, 0, 0)?.borderTop).toBe(RED);
    // The line the cell drew itself stays, even where the derived line has just moved away
    expect(formatAt(table, 0, 0)?.borderBottom).toBe(RED);
    expect(formatAt(table, 0, 1)?.borderBottom).toBe(INNER);
    expect(formatAt(table, 2, 1)?.borderBottom).toBe(OUTER);
  });

  it("keeps the fill and the vertical alignment a new cell inherited", () => {
    const props =
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>' +
      '<w:vAlign w:val="center"/>';
    const doc = tableDoc(
      2,
      rowXml(cellXml("a1", { props }), cellXml("b1")),
      rowXml(cellXml("a2"), cellXml("b2"))
    );
    const table = firstTable(edit(doc, [0, 0], addRowAfter));

    expect(formatAt(table, 1, 0)).toEqual({
      borderTop: INNER,
      borderBottom: INNER,
      borderLeft: OUTER,
      borderRight: INNER,
      background: "#FFFF00",
      verticalAlign: "center",
    });
  });

  it("wins over inherited table lines on all four shared edges", () => {
    const edited = runCommand(
      stateInCell(gridDoc(), 1, 1),
      setCellBorderColor("#FF0000")
    ).doc;
    const table = firstTable(edited);
    const red = "0.5pt solid #FF0000";

    expect(formatAt(table, 1, 1)).toMatchObject({
      borderTop: red,
      borderBottom: red,
      borderLeft: red,
      borderRight: red,
    });
    expect(formatAt(table, 0, 1)?.borderBottom).toBe("none");
    expect(formatAt(table, 1, 0)?.borderRight).toBe("none");
    expect(formatAt(table, 1, 2)?.borderLeft).toBe("none");
    expect(formatAt(table, 2, 1)?.borderTop).toBe("none");

    const map = TableMap.get(table);
    for (const [row, col] of [
      [0, 1],
      [1, 0],
      [1, 2],
      [2, 1],
    ] as const) {
      expect(
        table.nodeAt(map.map[row * map.width + col])?.attrs.tcPr
      ).not.toContain("FF0000");
    }
  });

  it("does not spread a direct color across the whole side of a merged neighbour", () => {
    const doc = tableDoc(
      2,
      rowXml(cellXml("wide", { span: 2 })),
      rowXml(cellXml("left"), cellXml("right"))
    );
    const table = firstTable(
      runCommand(stateInCell(doc, 1, 0), setCellBorderColor("#FF0000")).doc
    );

    expect(formatAt(table, 0, 0)?.borderBottom).toBe("none");
    expect(formatAt(table, 1, 0)?.borderTop).toBe("0.5pt solid #FF0000");
    expect(formatAt(table, 1, 1)?.borderTop).toBe(INNER);
  });
});

describe("an edit that leaves the grid as it was", () => {
  it("touches no cell at all", () => {
    const doc = gridDoc();
    const state = createEditorState(doc);
    const typed = state.apply(
      state.tr.insertText("edit", 1 + TableMap.get(firstTable(doc)).map[0] + 2)
    );

    expect(typed.doc.textContent).toContain("edit");
    // The cells of the rows below were not rewritten, so they are the very same nodes
    expect(firstTable(typed.doc).child(2)).toBe(firstTable(doc).child(2));
  });
});
