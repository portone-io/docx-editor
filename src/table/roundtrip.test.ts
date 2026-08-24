// @vitest-environment jsdom
/**
 * Guards the seam where table manipulation meets writing the file.
 *
 * A table whose rows or columns were edited has to still be a table once it is exported
 * and opened again.
 * A table whose grid and cell counts disagree falls back to a preserved block when it is
 * reopened, so "it is still a table after reopening" checks the grid invariant directly.
 */

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import {
  decode,
  fixtureNames,
  readFixture,
  surroundings,
} from "../__testing__/docx";
import { runCommand } from "../__testing__/editing";
import { exportDocx } from "../docx/exportDocx";
import { importDocx } from "../docx/importDocx";
import {
  A4_PORTRAIT,
  bodyWidth,
  type PageGeometry,
} from "../docx/pageGeometry";
import { serializeTable } from "../docx/serializeTable";
import type { SessionStore } from "../docx/session";
import { createEditorState } from "../editor/createEditor";
import { toCellFormat, toTableWidth } from "../model/format";
import { setCellBackground, setCellBorders } from "./cellFormatting";
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  type TableCommand,
} from "./commands";
import { buildMergeCellsTransaction, canMergeCells } from "./merge";
import { buildResizeColumnTransaction, maxGridTotal } from "./resize";

const MERGED = "kitchen-sink.docx";

/** A table with no merged cells, so an edit to one cell touches one `w:tc` */
const UNMERGED = "east-asian.docx";

/** A table stored wider than the body of the paper its document is written on */
const WIDER_THAN_BODY = "size-fallback.docx";

interface FoundTable {
  table: PMNode;
  /** The position of the first cell inside the table */
  start: number;
  /** Which body block the table is */
  index: number;
}

function firstTable(doc: PMNode): FoundTable {
  let found: FoundTable | null = null;
  doc.forEach((block, offset, index) => {
    if (!found && block.type.name === "table") {
      found = { table: block, start: offset + 1, index };
    }
  });
  if (!found) throw new Error("no table");
  return found;
}

/** A state with the cursor placed at one grid coordinate of the first table */
function stateInCell(doc: PMNode, row: number, col: number): EditorState {
  const { table, start } = firstTable(doc);
  const map = TableMap.get(table);
  const cellPos = map.map[row * map.width + col];
  const state = createEditorState(doc);
  return state.apply(
    state.tr.setSelection(
      TextSelection.near(state.doc.resolve(start + cellPos + 1))
    )
  );
}

function documentXmlOf(doc: PMNode, session: SessionStore): string {
  return decode(unzipSync(exportDocx(doc, session))[session.mainPartPath]);
}

/** Exports the edited table, opens it again, and returns the first table */
function reopenFirstTable(doc: PMNode, session: SessionStore): PMNode {
  const { doc: again } = importDocx(exportDocx(doc, session));
  return firstTable(again).table;
}

/** The list of grid column widths. Empty for a table that has no grid */
function gridColsOf(table: PMNode): number[] {
  const value: unknown = table.attrs.gridCols;
  return Array.isArray(value)
    ? value.filter((entry): entry is number => typeof entry === "number")
    : [];
}

function gridTotalOf(table: PMNode): number {
  return gridColsOf(table).reduce((sum, width) => sum + width, 0);
}

/**
 * Each command starts from a cell where that operation is allowed.
 * The first cell of the kitchen sink table is a merged cell covering the whole height of the
 * table, so no row can be deleted from there.
 */
const STRUCTURAL_COMMANDS: ReadonlyArray<
  [string, TableCommand, [number, number]]
> = [
  ["addRowBefore", addRowBefore, [0, 0]],
  ["addRowAfter", addRowAfter, [0, 0]],
  ["addColumnBefore", addColumnBefore, [0, 0]],
  ["addColumnAfter", addColumnAfter, [0, 0]],
  ["deleteColumn", deleteColumn, [0, 0]],
  ["deleteRow", deleteRow, [2, 3]],
];

describe("exporting a table whose rows and columns were edited", () => {
  it.each(STRUCTURAL_COMMANDS)(
    "kitchen sink: still a table on reopening after %s",
    (_name, command, [row, col]) => {
      const { doc, session } = importDocx(readFixture(MERGED));
      const edited = runCommand(stateInCell(doc, row, col), command);
      const reopened = reopenFirstTable(edited.doc, session);

      const gridCols = gridColsOf(reopened);
      expect(gridCols.length).toBeGreaterThan(0);
      expect(gridCols.length).toBe(TableMap.get(reopened).width);
    }
  );

  it("does not delete a row from a merged cell that covers the whole table height", () => {
    const { doc } = importDocx(readFixture(MERGED));
    expect(deleteRow(stateInCell(doc, 0, 0))).toBe(false);
  });

  it.each(fixtureNames)("%s: adding a row grows it by one row", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    const before = firstTable(doc).table.childCount;
    const edited = runCommand(stateInCell(doc, 0, 0), addRowAfter);

    expect(reopenFirstTable(edited.doc, session).childCount).toBe(before + 1);
  });

  it.each(fixtureNames)(
    "%s: adding a column grows the grid along with it",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const before = firstTable(doc).table;
      const edited = runCommand(stateInCell(doc, 0, 0), addColumnAfter);
      const reopened = reopenFirstTable(edited.doc, session);

      expect(TableMap.get(reopened).width).toBe(TableMap.get(before).width + 1);
      expect(gridColsOf(reopened).length).toBe(gridColsOf(before).length + 1);
    }
  );

  it.each(fixtureNames)(
    "%s: adding a column does not widen the table",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const before = firstTable(doc).table;
      const edited = runCommand(stateInCell(doc, 0, 0), addColumnAfter);
      const reopened = reopenFirstTable(edited.doc, session);

      // Only a table whose width is a share of the page grows its grid. Even then, the
      // table stays inside the page.
      const share = toTableWidth(before.attrs.tblW)?.type === "pct";
      const grown = share ? gridColsOf(before)[0] : 0;
      expect(gridTotalOf(reopened)).toBe(gridTotalOf(before) + grown);
      expect(reopened.attrs.tblW).toEqual(before.attrs.tblW);
    }
  );
});

/** The document with one column boundary of the first table dragged */
function resizeFirstTable(
  doc: PMNode,
  col: number,
  width: number,
  geometry: PageGeometry = A4_PORTRAIT
): PMNode {
  const state = createEditorState(doc);
  const tr = buildResizeColumnTransaction(
    state,
    { tablePos: firstTable(doc).start - 1, col, width },
    geometry
  );
  if (!tr) throw new Error("could not move the boundary");
  return state.apply(tr).doc;
}

/** The widths recorded by the cells of the first row */
function firstRowWidths(table: PMNode): unknown[] {
  const widths: unknown[] = [];
  table.child(0).forEach((cell) => widths.push(cell.attrs.tcW));
  return widths;
}

describe("exporting a table whose column widths were dragged", () => {
  it.each(fixtureNames)(
    "%s: widths moved by an inner boundary stay the same on reopening",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const before = firstTable(doc).table;
      const gridCols = gridColsOf(before);
      // Push the first boundary to the right by a third of the right-hand column
      const edited = resizeFirstTable(
        doc,
        0,
        gridCols[0] + Math.round(gridCols[1] / 3),
        session.geometry
      );
      const editedTable = firstTable(edited).table;
      const reopened = reopenFirstTable(edited, session);

      expect(gridColsOf(editedTable)).not.toEqual(gridCols);
      expect(gridColsOf(reopened)).toEqual(gridColsOf(editedTable));
      expect(firstRowWidths(reopened)).toEqual(firstRowWidths(editedTable));
      // An inner boundary, so the table neither widens nor narrows
      expect(gridTotalOf(reopened)).toBe(gridTotalOf(before));
      expect(reopened.attrs.tblW).toEqual(before.attrs.tblW);
    }
  );

  it.each(fixtureNames)(
    "%s: a table widened by the end boundary reopens unchanged, inside the body width",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const gridCols = gridColsOf(firstTable(doc).table);
      const last = gridCols.length - 1;
      // The body width of the paper this document names, which is where the table stops
      const body = bodyWidth(session.geometry).twips;
      // Halve the last column first: a table already as wide as its body cannot be widened,
      // and one of the fixtures is
      const narrowed = resizeFirstTable(
        doc,
        last,
        Math.round(gridCols[last] / 2),
        session.geometry
      );
      const before = firstTable(narrowed).table;
      // Deliberately pass a value past the body width. The table should widen but stop
      // at the body width.
      const edited = resizeFirstTable(
        narrowed,
        last,
        body * 2,
        session.geometry
      );
      const editedTable = firstTable(edited).table;
      const reopened = reopenFirstTable(edited, session);

      expect(gridTotalOf(editedTable)).toBeGreaterThan(gridTotalOf(before));
      // The table stops at the room the body leaves it, which is the body width less the
      // table's indent. A table indented outward (a negative `w:tblInd`) has that much more
      expect(gridTotalOf(editedTable)).toBe(
        maxGridTotal(before, session.geometry)
      );
      expect(gridTotalOf(editedTable)).toBeLessThanOrEqual(body + 108);
      expect(gridColsOf(reopened)).toEqual(gridColsOf(editedTable));
      expect(reopened.attrs.tblW).toEqual(editedTable.attrs.tblW);
    }
  );

  /**
   * The ceiling is the document's own body width, so a table stored wider than the paper it
   * is written on can only be narrowed. `size-fallback.docx` holds one: 9084 dxa of grid on
   * a body of 9026.
   */
  it("does not widen a table already wider than the document's body", () => {
    const { doc, session } = importDocx(readFixture(WIDER_THAN_BODY));
    const table = firstTable(doc).table;
    const last = gridColsOf(table).length - 1;
    const body = bodyWidth(session.geometry).twips;

    expect(gridTotalOf(table)).toBeGreaterThan(body);
    expect(maxGridTotal(table, session.geometry)).toBe(gridTotalOf(table));
    expect(() =>
      resizeFirstTable(doc, last, body * 2, session.geometry)
    ).toThrow();
  });
});

interface Mergeable {
  state: EditorState;
  row: number;
  columns: [number, number];
}

/** The first place where two side-by-side cells can be merged */
function firstMergeable(doc: PMNode): Mergeable | null {
  const { table, start } = firstTable(doc);
  const map = TableMap.get(table);
  const state = createEditorState(doc);

  for (let row = 0; row < map.height; row++) {
    for (let col = 0; col + 1 < map.width; col++) {
      const left = map.map[row * map.width + col];
      const right = map.map[row * map.width + col + 1];
      if (left === right) continue;
      const candidate = state.apply(
        state.tr.setSelection(
          CellSelection.create(doc, start + left, start + right)
        )
      );
      if (canMergeCells(candidate)) {
        return { state: candidate, row, columns: [col, col + 1] };
      }
    }
  }
  return null;
}

describe("exporting a merged table", () => {
  it("kitchen sink: the merged cell's width goes out as the sum of the grid columns it covers", () => {
    const { doc, session } = importDocx(readFixture(MERGED));
    const mergeable = firstMergeable(doc);
    if (!mergeable) throw new Error("no mergeable cells");

    const tr = buildMergeCellsTransaction(mergeable.state);
    if (!tr) throw new Error("the merge was refused");
    const merged = mergeable.state.apply(tr);

    const { table } = firstTable(merged.doc);
    const [left, right] = mergeable.columns;
    const gridCols = gridColsOf(table);
    const expected = gridCols[left] + gridCols[right];

    const map = TableMap.get(table);
    const mergedCell = table.nodeAt(map.map[mergeable.row * map.width + left]);
    expect(mergedCell?.attrs).toMatchObject({
      colspan: 2,
      tcW: { type: "dxa", twips: expected },
    });

    // The exported XML records that same width, and the grid still lines up on reopening
    expect(documentXmlOf(merged.doc, session)).toContain(
      `<w:tcW w:w="${expected}" w:type="dxa"/>`
    );
    expect(TableMap.get(reopenFirstTable(merged.doc, session)).width).toBe(
      map.width
    );
  });
});

/**
 * The one region where two exports of the same table differ.
 * Everything before `start` and everything after the two ends is identical text.
 */
function divergence(
  before: string,
  after: string
): { start: number; beforeEnd: number; afterEnd: number } {
  const limit = Math.min(before.length, after.length);
  let start = 0;
  while (start < limit && before[start] === after[start]) start += 1;
  let tail = 0;
  while (
    tail < before.length - start &&
    tail < after.length - start &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]
  ) {
    tail += 1;
  }
  return {
    start,
    beforeEnd: before.length - tail,
    afterEnd: after.length - tail,
  };
}

/** Where the cell that encloses this spot opens. A cell may or may not carry attributes */
function cellOpensAt(xml: string, at: number): number {
  return Math.max(xml.lastIndexOf("<w:tc>", at), xml.lastIndexOf("<w:tc ", at));
}

/** The display values the cell at one grid coordinate of the first table holds */
function cellFormatAt(table: PMNode, row: number, col: number) {
  const map = TableMap.get(table);
  const cell = table.nodeAt(map.map[row * map.width + col]);
  return toCellFormat(cell?.attrs.format);
}

describe("exporting a table whose cell borders and fill were edited", () => {
  it.each(fixtureNames)(
    "%s: the borders and the fill are there again on reopening",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const bordered = runCommand(
        stateInCell(doc, 0, 0),
        setCellBorders("all")
      );
      const filled = runCommand(bordered, setCellBackground("#ffff00"));
      const reopened = reopenFirstTable(filled.doc, session);

      expect(cellFormatAt(reopened, 0, 0)).toMatchObject({
        background: "#FFFF00",
        borderTop: "0.5pt solid #000000",
        borderBottom: "0.5pt solid #000000",
        borderLeft: "0.5pt solid #000000",
        borderRight: "0.5pt solid #000000",
      });
      // The grid still lines up, so the table came back as a table
      const gridCols = gridColsOf(reopened);
      expect(gridCols.length).toBe(TableMap.get(reopened).width);
    }
  );

  it(`${UNMERGED}: nothing but the edited cell's own XML changes`, () => {
    const { doc } = importDocx(readFixture(UNMERGED));
    const before = serializeTable(firstTable(doc).table);
    const edited = runCommand(stateInCell(doc, 1, 1), setCellBorders("all"));
    const after = serializeTable(firstTable(edited.doc).table);
    const { start, beforeEnd, afterEnd } = divergence(before, after);

    expect(after).not.toBe(before);
    // The change opens inside one cell and closes before that cell does
    expect(cellOpensAt(before, start)).toBeGreaterThan(
      before.lastIndexOf("</w:tc>", start)
    );
    expect(before.slice(start, beforeEnd)).not.toContain("</w:tc>");
    expect(after.slice(start, afterEnd)).not.toContain("</w:tc>");
    expect(after.slice(start, afterEnd)).toContain('<w:top w:val="single"');
  });
});

/** Every edit that reaches into a table and nothing else */
const TABLE_EDITS: ReadonlyArray<[string, (doc: PMNode) => PMNode]> = [
  [
    "a row is added",
    (doc) => runCommand(stateInCell(doc, 0, 0), addRowAfter).doc,
  ],
  [
    "a column boundary is dragged",
    (doc) => {
      const gridCols = gridColsOf(firstTable(doc).table);
      return resizeFirstTable(
        doc,
        0,
        gridCols[0] + Math.round(gridCols[1] / 3)
      );
    },
  ],
  [
    "a cell is filled",
    (doc) =>
      runCommand(stateInCell(doc, 0, 0), setCellBackground("#ffff00")).doc,
  ],
];

describe.each(TABLE_EDITS)("once %s", (_edit, edit) => {
  it.each(fixtureNames)(
    "%s: the blocks outside the table stay as the original bytes",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const { index } = firstTable(doc);
      const documentXml = documentXmlOf(edit(doc), session);
      const { head, tail } = surroundings(session, index);

      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);
    }
  );
});
