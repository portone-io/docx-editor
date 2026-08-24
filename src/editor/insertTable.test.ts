// @vitest-environment jsdom
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { fixTables } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import {
  decode,
  LETTER_FIXTURE,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
} from "../__testing__/docx";
import { posOfText } from "../__testing__/editing";
import { exportDocx } from "../docx/exportDocx";
import { importDocx } from "../docx/importDocx";
import { A4_BODY_WIDTH } from "../docx/pageGeometry";
import type { SessionStore } from "../docx/session";
import { createTableNode } from "../docx/tableTemplate";
import { type TableWidth, toGridCols, toTableWidth } from "../model/format";
import { createEditorState } from "./createEditor";
import { canInsertTable, insertTable } from "./insertTable";

const BODY = '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>';

/** A committed document on A4, to set beside the one on Letter */
const A4_FIXTURE = "kitchen-sink.docx";

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const TABLE_BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cellXml("Left")}${cellXml("Right")}</w:tr>` +
  "</w:tbl>";

/** An editing state with the caret on a piece of text, plus the session used to write that document back out */
function openWithCaret(
  body: string,
  needle: string
): { state: EditorState; session: SessionStore } {
  const { doc, session } = importDocx(makeDocx(body));
  const state = createEditorState(doc);
  return {
    state: state.apply(
      state.tr.setSelection(TextSelection.create(doc, posOfText(doc, needle)))
    ),
    session,
  };
}

/** Runs the command and returns the next state. Null when there is nothing to do */
function run(
  state: EditorState,
  rows: number,
  cols: number
): EditorState | null {
  let next: EditorState | null = null;
  const handled = insertTable({ rows, columns: cols })(state, (tr) => {
    next = state.apply(tr);
  });
  return handled ? next : null;
}

function insertedTable(state: EditorState, rows: number, cols: number) {
  const next = run(state, rows, cols);
  if (!next) throw new Error("could not insert the table");
  return next;
}

function cellsOf(row: PMNode): PMNode[] {
  const cells: PMNode[] = [];
  row.forEach((cell) => cells.push(cell));
  return cells;
}

describe("creating a new table", () => {
  it("divides the column widths evenly so that their sum equals the body width", () => {
    for (const cols of [1, 2, 3, 4, 5, 6, 7]) {
      const gridCols = toGridCols(createTableNode(2, cols).attrs.gridCols);
      expect(gridCols).toHaveLength(cols);
      expect(gridCols.reduce((sum, w) => sum + w, 0)).toBe(A4_BODY_WIDTH.twips);
      // Since the widths are divided evenly, the widest and narrowest columns differ by at most 1 twip
      expect(Math.max(...gridCols) - Math.min(...gridCols)).toBeLessThanOrEqual(
        1
      );
    }
  });

  it("writes down on each cell the width of its own grid column", () => {
    const table = createTableNode(2, 3);
    const gridCols = toGridCols(table.attrs.gridCols);
    expect(table.childCount).toBe(2);

    table.forEach((row) => {
      const widths = cellsOf(row).map((cell) => toTableWidth(cell.attrs.tcW));
      expect(widths).toEqual(gridCols.map((twips) => ({ type: "dxa", twips })));
    });
  });

  it("the table width is the body width and every cell shows a black 0.5pt border", () => {
    expect(toTableWidth(createTableNode(1, 2).attrs.tblW)).toEqual({
      type: "dxa",
      twips: A4_BODY_WIDTH.twips,
    });
    // The same line runs around the table and between its cells, wherever a cell sits in the grid
    const line = "0.5pt solid #000000";
    createTableNode(3, 3).forEach((row) => {
      for (const cell of cellsOf(row)) {
        expect(cell.attrs.format).toMatchObject({
          borderTop: line,
          borderLeft: line,
          borderBottom: line,
          borderRight: line,
        });
      }
    });
  });

  it("every cell holds a single empty paragraph", () => {
    const table = createTableNode(2, 2);
    table.forEach((row) => {
      for (const cell of cellsOf(row)) {
        expect(cell.childCount).toBe(1);
        expect(cell.child(0).type.name).toBe("paragraph");
        expect(cell.child(0).content.size).toBe(0);
      }
    });
  });
});

describe("the insert table command", () => {
  it("inserts after the paragraph holding the caret and moves the caret into the first cell", () => {
    const { state } = openWithCaret(BODY, "body");
    expect(canInsertTable(state)).toBe(true);

    const next = insertedTable(state, 2, 3);
    expect(next.doc.childCount).toBe(3);
    expect(next.doc.child(0).textContent).toBe("body");
    expect(next.doc.child(1).type.name).toBe("table");
    expect(next.doc.child(2).type.name).toBe("paragraph");
    expect(next.doc.child(2).content.size).toBe(0);

    // The caret is inside the empty paragraph of the first cell (doc > table > tr > tc > p)
    const $from = next.selection.$from;
    expect($from.depth).toBe(4);
    expect($from.node(3)).toBe(next.doc.child(1).child(0).child(0));
  });

  it("does not insert inside a table", () => {
    const { state } = openWithCaret(TABLE_BODY, "Left");
    expect(canInsertTable(state)).toBe(false);
    expect(insertTable({ rows: 2, columns: 2 })(state)).toBe(false);
  });

  it("does nothing when the row or column count makes no sense", () => {
    const { state } = openWithCaret(BODY, "body");
    expect(insertTable({ rows: 0, columns: 2 })(state)).toBe(false);
    expect(insertTable({ rows: 2, columns: 0 })(state)).toBe(false);
    expect(insertTable({ rows: 1.5, columns: 2 })(state)).toBe(false);
    expect(insertTable({ rows: 2, columns: 1000 })(state)).toBe(false);
  });

  it("the inserted table is a grid with nothing to fix", () => {
    const { state } = openWithCaret(BODY, "body");
    expect(fixTables(insertedTable(state, 3, 4))).toBeUndefined();
  });

  it("fits the table to the paper the document names, not to A4", () => {
    const { doc, session } = importDocx(makeDocx(BODY + LETTER_SECT_PR));
    const opened = createEditorState(doc, { geometry: session.geometry });
    const state = opened.apply(
      opened.tr.setSelection(TextSelection.create(doc, posOfText(doc, "body")))
    );

    const table = insertedTable(state, 2, 3).doc.child(1);
    // Letter with an inch of margin leaves 9360 dxa of body, where A4 leaves 9412
    expect(toTableWidth(table.attrs.tblW)).toEqual({
      type: "dxa",
      twips: 9360,
    });
    const gridCols = toGridCols(table.attrs.gridCols);
    expect(gridCols.reduce((sum, width) => sum + width, 0)).toBe(9360);
    expect(gridCols).toEqual([3120, 3120, 3120]);
    expect(gridCols.reduce((sum, width) => sum + width, 0)).not.toBe(
      A4_BODY_WIDTH.twips
    );
  });

  /**
   * The same insertion on two committed documents. Neither number is the A4 fallback's 9412:
   * each is the body of the paper its own document is written on.
   */
  it("gives a Letter document and an A4 document tables of different widths", () => {
    const widthIn = (name: string): TableWidth | null => {
      const { doc, session } = importDocx(readFixture(name));
      const opened = createEditorState(doc, { geometry: session.geometry });
      const state = opened.apply(
        opened.tr.setSelection(TextSelection.near(opened.doc.resolve(1)))
      );
      const inserted = insertedTable(state, 2, 2).doc;
      const tables: PMNode[] = [];
      inserted.forEach((block) => {
        // The new table is the one with no original fragment behind it
        if (block.type.name === "table" && block.attrs.srcId === null) {
          tables.push(block);
        }
      });
      expect(tables).toHaveLength(1);
      return toTableWidth(tables[0].attrs.tblW);
    };

    expect(widthIn(LETTER_FIXTURE)).toEqual({ type: "dxa", twips: 9360 });
    expect(widthIn(A4_FIXTURE)).toEqual({ type: "dxa", twips: 9026 });
    // Neither is the A4 fallback the editor used to fit every table to
    expect(A4_BODY_WIDTH.twips).toBe(9412);
  });
});

describe("exporting an inserted table and reopening it", () => {
  it("the grid and the formatting fragments come back unchanged", () => {
    const { state, session } = openWithCaret(BODY, "body");
    const next = insertedTable(state, 2, 3);
    const inserted = next.doc.child(1);

    const reopened = importDocx(exportDocx(next.doc, session)).doc;
    expect(reopened.childCount).toBe(3);

    const table = reopened.child(1);
    expect(table.type.name).toBe("table");
    expect(toGridCols(table.attrs.gridCols)).toEqual(
      toGridCols(inserted.attrs.gridCols)
    );
    // The reopened table differs only in its original-fragment id; everything else is as inserted
    expect(table.attrs.tblPr).toBe(inserted.attrs.tblPr);
    expect(table.attrs.format).toEqual(inserted.attrs.format);
    expect(table.childCount).toBe(2);
    expect(reopened.child(2).type.name).toBe("paragraph");
    expect(reopened.child(2).content.size).toBe(0);
    expect(table.child(0).childCount).toBe(3);
    expect(table.child(0).child(0).attrs.tcPr).toBe(
      inserted.child(0).child(0).attrs.tcPr
    );
    expect(table.child(0).child(0).attrs.format).toEqual(
      inserted.child(0).child(0).attrs.format
    );
  });

  it("leaves the body paragraph at its original bytes and writes out only the table anew", () => {
    const { state, session } = openWithCaret(BODY, "body");
    const next = insertedTable(state, 1, 2);

    const documentXml = decode(
      unzipSync(exportDocx(next.doc, session))[session.mainPartPath]
    );
    expect(documentXml).toContain(session.blocks[0].xml + "<w:tbl>");
    expect(documentXml).toContain(
      `<w:tblW w:w="${A4_BODY_WIDTH.twips}" w:type="dxa"/>`
    );
    // Only by pinning the widths to the grid does Word render them at the same widths as the screen
    expect(documentXml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(documentXml).toContain("<w:gridCol");
  });
});
