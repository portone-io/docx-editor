// @vitest-environment node
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { fixTables, TableMap } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import {
  cell,
  cellsOf,
  cellWithText,
  dxa,
  firstTable,
  mergedContractDoc,
  posOfText,
  row,
  rowsOf,
  rowWith,
  sampleContractDoc,
  schema,
  stateWithCursorIn,
  tableDoc,
  uniformDoc,
} from "./__testing__/tables";
import {
  buildAddRowAfterTransaction,
  buildAddRowBeforeTransaction,
  buildDeleteRowTransaction,
} from "./rows";

function addRowAfterCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildAddRowAfterTransaction(state);
  if (!tr) throw new Error("not in table");
  return state.apply(tr);
}

function addRowBeforeCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildAddRowBeforeTransaction(state);
  if (!tr) throw new Error("not in table");
  return state.apply(tr);
}

function deleteRowOfCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildDeleteRowTransaction(state);
  if (!tr) throw new Error("delete refused");
  return state.apply(tr);
}

/**
 * For comparison. The wrong approach, cloning a row wholesale without consulting the grid
 */
function naiveAddRowAfterCellWith(doc: PMNode, needle: string): EditorState {
  const state = EditorState.create({ doc, schema });
  const $cursor = doc.resolve(posOfText(doc, needle));
  let rowNode: PMNode | null = null;
  let before = -1;
  for (let depth = $cursor.depth; depth > 0; depth--) {
    if ($cursor.node(depth).type.spec.tableRole === "row") {
      rowNode = $cursor.node(depth);
      before = $cursor.before(depth);
      break;
    }
  }
  if (!rowNode) throw new Error("no row");
  const cells: PMNode[] = [];
  rowNode.forEach((current) => {
    const filled = current.type.createAndFill(current.attrs);
    if (filled) cells.push(filled);
  });
  const newRow = rowNode.type.create(rowNode.attrs, cells);
  const tr = state.tr.insert(before + rowNode.nodeSize, newRow);
  return state.apply(tr);
}

describe("buildAddRowAfterTransaction - merged cell table", () => {
  it("mid-span insert (after Address) extends the covering rowspan and needs no repair", () => {
    const state = addRowAfterCellWith(mergedContractDoc(), "Address");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(4);
    expect(map.height).toBe(4);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(4);
    expect(cellWithText(table, "Contact")).toBeDefined();
  });

  it("inherits tcW/colspan into the new cells (empty content)", () => {
    const state = addRowAfterCellWith(mergedContractDoc(), "Address");
    const { table } = firstTable(state.doc);
    const newCells = cellsOf(rowsOf(table)[2]);

    expect(newCells.length).toBe(2);
    expect(newCells[0].attrs.tcW).toEqual(dxa(1500));
    expect(newCells[1].attrs.colspan).toBe(2);
    expect(newCells[1].attrs.tcW).toEqual(dxa(4000));
    expect(newCells.every((current) => current.textContent === "")).toBe(true);
  });

  it("copies the listed format attrs and recomputes the spans", () => {
    const doc = tableDoc([
      row(
        cell("Head", {
          rowspan: 2,
          colwidth: [120],
          tcAttrs: ' w:id="1"',
          tcPr: '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>',
          tcW: dxa(1500),
          format: { background: "#ffff00" },
        }),
        cell("Value", { tcW: dxa(2000) })
      ),
      row(cell("Value2", { tcW: dxa(2000) })),
    ]);
    const state = addRowAfterCellWith(doc, "Value2");
    const { table } = firstTable(state.doc);
    const newCell = cellsOf(rowsOf(table)[2])[0];

    expect(newCell.attrs.tcAttrs).toBe(' w:id="1"');
    expect(newCell.attrs.tcPr).toBe(
      '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>'
    );
    expect(newCell.attrs.tcW).toEqual(dxa(1500));
    expect(newCell.attrs.format).toEqual({ background: "#ffff00" });
    // The new cell occupies the same grid columns as the reference cell, so its on-screen
    // width is unchanged too.
    expect(newCell.attrs.colwidth).toEqual([120]);
    // The merge counts are not inherited; they are counted again from the grid.
    expect(newCell.attrs.rowspan).toBe(1);
    expect(newCell.attrs.colspan).toBe(1);
  });

  it("inherits the row format from the reference row", () => {
    const doc = tableDoc([
      rowWith(
        {
          trAttrs: ' w:rsidR="00A"',
          trPr: '<w:trPr><w:trHeight w:val="400"/></w:trPr>',
          format: { heightPt: 20 },
        },
        cell("A", { tcW: dxa(1000) })
      ),
    ]);
    const state = addRowAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);
    const newRow = rowsOf(table)[1];

    expect(newRow.attrs.trAttrs).toBe(' w:rsidR="00A"');
    expect(newRow.attrs.trPr).toBe(
      '<w:trPr><w:trHeight w:val="400"/></w:trPr>'
    );
    expect(newRow.attrs.format).toEqual({ heightPt: 20 });
  });

  it("takes neither the content control of the reference cell nor the row's property exceptions along", () => {
    const doc = tableDoc([
      rowWith(
        { tblPrEx: "<w:tblPrEx><w:tblBorders/></w:tblPrEx>" },
        cell("A", { tcW: dxa(1000), sdtPrefix: "<w:sdt><w:sdtPr/>" })
      ),
    ]);
    const state = addRowAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);
    const newRow = rowsOf(table)[1];

    expect(newRow.attrs.tblPrEx).toBeNull();
    expect(cellsOf(newRow)[0].attrs.sdtPrefix).toBeNull();
  });

  it("last covered row insert (after Contact) stays valid, rowspan unchanged", () => {
    const state = addRowAfterCellWith(mergedContractDoc(), "Contact");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(4);
    expect(map.height).toBe(4);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(3);
  });

  it("REGRESSION: naive clone-the-row approach produces a malformed grid needing repair", () => {
    const state = naiveAddRowAfterCellWith(mergedContractDoc(), "Address");

    expect(fixTables(state)).toBeDefined();
    const { table } = firstTable(state.doc);
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(3);
  });
});

describe("buildAddRowBeforeTransaction", () => {
  it("inserts above the cursor row and keeps the grid valid", () => {
    const state = addRowBeforeCellWith(uniformDoc(), "Requirements");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(3);
    expect(map.height).toBe(4);
    expect(fixTables(state)).toBeUndefined();
    expect(rowsOf(table)[1].textContent).toBe("");
    expect(rowsOf(table)[2].textContent).toContain("Requirements");
  });

  it("inserting above the first row of a merged table extends the rowspan", () => {
    const state = addRowBeforeCellWith(mergedContractDoc(), "Address");
    const { table } = firstTable(state.doc);

    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(4);
    expect(rowsOf(table)[1].textContent).toBe("");
  });
});

describe("buildAddRowAfterTransaction - real contract (doubly-overlapping rowspan)", () => {
  it("insert after Contact (covered by BOTH Parties & Client) extends both rowspans, no shift/phantom", () => {
    const state = addRowAfterCellWith(sampleContractDoc(), "555-0100");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(5);
    expect(map.height).toBe(6);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "Parties")?.attrs.rowspan).toBe(6);
    expect(cellWithText(table, "Client")?.attrs.rowspan).toBe(5);

    const rows = rowsOf(table);
    expect(rows[4].textContent).toContain("Manager");

    const newCells = cellsOf(rows[3]);
    expect(newCells.length).toBe(2);
    expect(newCells[0].attrs.colspan).toBe(1);
    expect(newCells[1].attrs.colspan).toBe(2);
    expect(newCells.every((current) => current.textContent === "")).toBe(true);
  });

  it("REGRESSION: naive clone leaves both rowspans unchanged -> malformed (repair needed)", () => {
    const state = naiveAddRowAfterCellWith(sampleContractDoc(), "555-0100");

    expect(fixTables(state)).toBeDefined();
    const { table } = firstTable(state.doc);
    expect(cellWithText(table, "Parties")?.attrs.rowspan).toBe(5);
    expect(cellWithText(table, "Client")?.attrs.rowspan).toBe(4);
  });
});

describe("buildAddRowAfterTransaction - uniform table", () => {
  it("insert after a middle row keeps 3 columns, inherits tcW, needs no repair", () => {
    const state = addRowAfterCellWith(uniformDoc(), "Requirements");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(3);
    expect(map.height).toBe(4);
    expect(fixTables(state)).toBeUndefined();

    const newCells = cellsOf(rowsOf(table)[2]);
    expect(newCells.length).toBe(3);
    expect(newCells.map((current) => current.attrs.tcW)).toEqual([
      dxa(3000),
      dxa(3000),
      dxa(3000),
    ]);
    expect(newCells.every((current) => current.textContent === "")).toBe(true);
  });
});

describe("buildAddRowAfterTransaction - colspan not at the row end", () => {
  // Advancing the grid coordinate and the child index separately puts them out of sync by
  // the merge span.
  // Once they are out of sync, decisions are made against the wrong reference cell and
  // the composition of the new row falls apart.
  const colspanFirstDoc = () =>
    tableDoc([
      row(
        cell("A", { colspan: 2, tcW: dxa(2000) }),
        cell("B", { tcW: dxa(1000) }),
        cell("C", { tcW: dxa(1000) })
      ),
      row(
        cell("D", { tcW: dxa(1000) }),
        cell("E", { tcW: dxa(1000) }),
        cell("F", { tcW: dxa(1000) }),
        cell("G", { tcW: dxa(1000) })
      ),
    ]);

  it("REGRESSION: reproduces every reference cell, not just the merged one", () => {
    const state = addRowAfterCellWith(colspanFirstDoc(), "A");
    const { table } = firstTable(state.doc);
    const newCells = cellsOf(rowsOf(table)[1]);

    // Out of sync, the result is [colspan 2, colspan 2]: A's formatting is copied twice
    // and the slots for B and C disappear.
    // The grid width still comes to 4, so `fixTables` does not catch it and the cell
    // composition is inspected directly.
    expect(newCells.map((current) => current.attrs.colspan)).toEqual([2, 1, 1]);
    expect(newCells.map((current) => current.attrs.tcW)).toEqual([
      dxa(2000),
      dxa(1000),
      dxa(1000),
    ]);
    expect(TableMap.get(table).width).toBe(4);
    expect(fixTables(state)).toBeUndefined();
  });

  it("REGRESSION: colspan and rowspan on the same cell still yields a full row", () => {
    const doc = tableDoc([
      row(
        cell("A", { colspan: 2, rowspan: 2, tcW: dxa(2000) }),
        cell("B", { tcW: dxa(1000) })
      ),
      row(cell("C", { tcW: dxa(1000) })),
      row(
        cell("D", { tcW: dxa(1000) }),
        cell("E", { tcW: dxa(1000) }),
        cell("F", { tcW: dxa(1000) })
      ),
    ]);
    const state = addRowBeforeCellWith(doc, "C");
    const { table } = firstTable(state.doc);

    // Out of sync, the new row ends up with zero cells and `fixTables` demands a repair.
    expect(rowsOf(table).map((current) => current.childCount)).toEqual([
      2, 1, 1, 3,
    ]);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "A")?.attrs.rowspan).toBe(3);
  });
});

/**
 * Pins down the current behaviour. The upstream `deleteRow` of `prosemirror-tables`
 * deletes every row the selection rectangle covers, and refuses when the rectangle spans
 * the full height of the table. Placing the cursor in a merged cell makes the rectangle
 * that whole cell, so the outcome branches three ways.
 */
describe("buildDeleteRowTransaction", () => {
  it("deletes only the cursor row and shrinks both covering rowspans", () => {
    const state = deleteRowOfCellWith(sampleContractDoc(), "555-0100");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.height).toBe(4);
    expect(map.width).toBe(5);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "Parties")?.attrs.rowspan).toBe(4);
    expect(cellWithText(table, "Client")?.attrs.rowspan).toBe(3);
    expect(cellWithText(table, "555-0100")).toBeUndefined();
  });

  it("shrinks a rowspan 3 cell to 2 when a covered row goes", () => {
    const state = deleteRowOfCellWith(mergedContractDoc(), "Address");
    const { table } = firstTable(state.doc);

    expect(TableMap.get(table).height).toBe(2);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(2);
  });

  it("deletes every row the merged cell spans (Client, rowspan 4 of 5)", () => {
    const state = deleteRowOfCellWith(sampleContractDoc(), "Client");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    // Four contract clauses go away together and only the last row remains. That is the
    // upstream contract.
    expect(map.height).toBe(1);
    expect(map.width).toBe(5);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "Parties")?.attrs.rowspan).toBe(1);
    expect(cellWithText(table, "Client")).toBeUndefined();
    expect(table.textContent).toBe("PartiesInstructorCompany[  ]");
  });

  it("refuses when the merged cell spans the whole table (Parties, rowspan 5 of 5)", () => {
    expect(
      buildDeleteRowTransaction(
        stateWithCursorIn(sampleContractDoc(), "Parties")
      )
    ).toBeNull();
  });

  it("leaves gridCols and the table width alone", () => {
    const doc = tableDoc(
      [
        row(cell("A", { tcW: dxa(1000) }), cell("B", { tcW: dxa(2000) })),
        row(cell("a", { tcW: dxa(1000) }), cell("b", { tcW: dxa(2000) })),
        row(cell("x", { tcW: dxa(1000) }), cell("y", { tcW: dxa(2000) })),
      ],
      { gridCols: [1000, 2000], tblW: dxa(3000) }
    );
    const state = deleteRowOfCellWith(doc, "a");
    const { table } = firstTable(state.doc);

    expect(TableMap.get(table).height).toBe(2);
    expect(table.attrs.gridCols).toEqual([1000, 2000]);
    expect(table.attrs.tblW).toEqual(dxa(3000));
  });

  it("returns null outside a table", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("Body")),
    ]);

    expect(
      buildDeleteRowTransaction(EditorState.create({ doc, schema }))
    ).toBeNull();
  });
});
