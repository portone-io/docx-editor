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
  pct,
  row,
  rowsOf,
  sampleContractDoc,
  schema,
  stateWithCursorIn,
  tableDoc,
} from "./__testing__/tables";
import {
  buildAddColumnAfterTransaction,
  buildAddColumnBeforeTransaction,
  buildDeleteColumnTransaction,
} from "./columns";

/**
 * Collects which grid column each of the cells just inserted landed on.
 *
 * Every cell of the fake tables has content, so "the first empty cell in each row" is
 * the new cell.
 * If the column insertion is right, all the values are the same single insertion point.
 * A table that is out of alignment still has uniform row widths, so `fixTables` does not
 * catch it and these coordinates are inspected directly.
 */
function gridColumnsOfEmptyCells(table: PMNode): number[] {
  const map = TableMap.get(table);
  const columns: number[] = [];
  for (let rowIndex = 0; rowIndex < map.height; rowIndex++) {
    for (let col = 0; col < map.width; col++) {
      const index = rowIndex * map.width + col;
      const pos = map.map[index];
      if (map.colCount(pos) !== col) continue;
      if (rowIndex > 0 && map.map[index - map.width] === pos) continue;
      if (table.nodeAt(pos)?.textContent === "") columns.push(col);
    }
  }
  return columns;
}

/** The sum of the grid column widths, which is the width the table actually occupies */
function gridSum(table: PMNode): number {
  const value: unknown = table.attrs.gridCols;
  const cols = Array.isArray(value) ? value : [];
  return cols.reduce<number>(
    (sum, width) => sum + (typeof width === "number" ? width : 0),
    0
  );
}

function addColumnAfterCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildAddColumnAfterTransaction(state);
  if (!tr) throw new Error("not in table");
  return state.apply(tr);
}

function addColumnBeforeCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildAddColumnBeforeTransaction(state);
  if (!tr) throw new Error("not in table");
  return state.apply(tr);
}

function deleteColumnOfCellWith(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildDeleteColumnTransaction(state);
  if (!tr) throw new Error("delete refused");
  return state.apply(tr);
}

/**
 * For comparison. The wrong approach, inserting by "after the nth cell of each row"
 * rather than by the grid
 */
function naiveAddColumnAt(doc: PMNode, childIndex: number): EditorState {
  const state = EditorState.create({ doc, schema });
  const { table, start } = firstTable(doc);
  const tr = state.tr;
  let rowStart = start;
  table.forEach((rowNode) => {
    let offset = rowStart + 1;
    for (let i = 0; i < Math.min(childIndex + 1, rowNode.childCount); i++) {
      offset += rowNode.child(i).nodeSize;
    }
    tr.insert(tr.mapping.map(offset), cell("", { tcW: dxa(100) }));
    rowStart += rowNode.nodeSize;
  });
  return state.apply(tr);
}

describe("buildAddColumnAfterTransaction - merged cell table", () => {
  it("puts the new cell on the same grid column in every row", () => {
    const state = addColumnAfterCellWith(mergedContractDoc(), "Company");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(5);
    expect(map.height).toBe(3);
    expect(fixTables(state)).toBeUndefined();
    expect(gridColumnsOfEmptyCells(table)).toEqual([2, 2, 2]);
  });

  it("REGRESSION: naive per-row child insert zigzags across columns and fixTables misses it", () => {
    const state = naiveAddColumnAt(mergedContractDoc(), 1);
    const { table } = firstTable(state.doc);

    expect(fixTables(state)).toBeUndefined();
    expect(gridColumnsOfEmptyCells(table)).toEqual([2, 4, 3]);
  });

  it("widens a crossing colspan cell instead of adding a cell inside it", () => {
    // Grid column 3 is right in the middle of the colspan-2 cells of the top two rows.
    const state = addColumnAfterCellWith(mergedContractDoc(), "555-0123");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(5);
    expect(fixTables(state)).toBeUndefined();
    expect(cellWithText(table, "Acme Corp")?.attrs.colspan).toBe(3);
    expect(cellWithText(table, "10 Main Street")?.attrs.colspan).toBe(3);
    expect(gridColumnsOfEmptyCells(table)).toEqual([3]);
    expect(cellsOf(rowsOf(table)[2]).length).toBe(4);
  });

  it("copies the listed format attrs and recomputes the spans", () => {
    const doc = tableDoc([
      row(
        cell("Head", {
          rowspan: 1,
          colwidth: [120],
          tcAttrs: ' w:id="1"',
          tcPr: '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>',
          tcW: dxa(1500),
          format: { background: "#ffff00" },
        }),
        cell("Value", { tcW: dxa(2000) })
      ),
    ]);
    const state = addColumnAfterCellWith(doc, "Head");
    const { table } = firstTable(state.doc);
    const newCell = cellsOf(rowsOf(table)[0])[1];

    expect(newCell.textContent).toBe("");
    expect(newCell.attrs.colspan).toBe(1);
    expect(newCell.attrs.rowspan).toBe(1);
    expect(newCell.attrs.tcAttrs).toBe(' w:id="1"');
    expect(newCell.attrs.tcPr).toBe(
      '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>'
    );
    expect(newCell.attrs.tcW).toEqual(dxa(1500));
    expect(newCell.attrs.format).toEqual({ background: "#ffff00" });
    expect(newCell.attrs.colwidth).toEqual([120]);
  });

  it("splits a colspan reference width instead of copying the merged total", () => {
    const doc = tableDoc([
      row(cell("Merged", { colspan: 2, tcW: dxa(4000) })),
      row(cell("Left", { tcW: dxa(2000) }), cell("Right", { tcW: dxa(2000) })),
    ]);
    const state = addColumnAfterCellWith(doc, "Right");
    const { table } = firstTable(state.doc);
    const appended = cellsOf(rowsOf(table)[0])[1];

    expect(fixTables(state)).toBeUndefined();
    expect(appended.textContent).toBe("");
    expect(appended.attrs.colspan).toBe(1);
    expect(appended.attrs.tcW).toEqual(dxa(2000));
  });

  it("gives no width to a new cell whose reference had none", () => {
    const doc = tableDoc([row(cell("A"), cell("B"))], {
      gridCols: [1000, 2000],
    });
    const state = addColumnAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);

    expect(cellsOf(rowsOf(table)[0])[1].attrs.tcW).toBeNull();
  });
});

describe("buildAddColumnAfterTransaction - doubly-overlapping rowspan", () => {
  it("adds one cell per grid row on the same column, no repair needed", () => {
    const state = addColumnAfterCellWith(sampleContractDoc(), "Company");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(6);
    expect(map.height).toBe(5);
    expect(fixTables(state)).toBeUndefined();
    expect(gridColumnsOfEmptyCells(table)).toEqual([3, 3, 3, 3, 3]);
    expect(cellWithText(table, "Parties")?.attrs.rowspan).toBe(5);
    expect(cellWithText(table, "Client")?.attrs.rowspan).toBe(4);
  });
});

describe("buildAddColumnBeforeTransaction", () => {
  it("inserting at the very left keeps the grid valid", () => {
    const state = addColumnBeforeCellWith(mergedContractDoc(), "PartyB");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(5);
    expect(fixTables(state)).toBeUndefined();
    expect(gridColumnsOfEmptyCells(table)).toEqual([0, 0, 0]);
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(3);
  });
});

describe("gridCols sync", () => {
  const gridDoc = () =>
    tableDoc(
      [
        row(
          cell("A", { tcW: dxa(1000) }),
          cell("B", { tcW: dxa(2000) }),
          cell("C", { tcW: dxa(3000) })
        ),
        row(
          cell("a", { tcW: dxa(1000) }),
          cell("bc", { colspan: 2, tcW: dxa(5000) })
        ),
      ],
      { gridCols: [1000, 2000, 3000] }
    );

  it("inserts one entry at the new column and reshares the same total", () => {
    const state = addColumnAfterCellWith(gridDoc(), "A");
    const { table } = firstTable(state.doc);

    // The 7000 left after slotting in the reference column's 1000 is brought back to
    // 6000 (a factor of 6/7). The order stays the same.
    expect(TableMap.get(table).width).toBe(4);
    expect(table.attrs.gridCols).toEqual([857, 857, 1714, 2572]);
  });

  it("appending at the right end also takes its width from the others", () => {
    const state = addColumnAfterCellWith(gridDoc(), "C");
    const { table } = firstTable(state.doc);

    // The 9000 left after slotting in the last column's 3000 is brought back to 6000 (a
    // factor of 2/3).
    expect(TableMap.get(table).width).toBe(4);
    expect(table.attrs.gridCols).toEqual([667, 1333, 2000, 2000]);
  });

  it("deleting a column removes exactly that entry and shrinks the crossing colspan", () => {
    const state = deleteColumnOfCellWith(gridDoc(), "B");
    const { table } = firstTable(state.doc);

    expect(TableMap.get(table).width).toBe(2);
    expect(fixTables(state)).toBeUndefined();
    expect(table.attrs.gridCols).toEqual([1000, 3000]);
    expect(cellWithText(table, "bc")?.attrs.colspan).toBe(1);
  });

  it("treats an empty grid as no grid", () => {
    const doc = tableDoc(
      [row(cell("A", { tcW: dxa(1000) }), cell("B", { tcW: dxa(2000) }))],
      { gridCols: [], tblW: dxa(3000) }
    );
    const state = addColumnAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);

    // No zero-width column is created in a table whose grid is unknown.
    expect(table.attrs.gridCols).toEqual([]);
    expect(table.attrs.tblW).toEqual(dxa(3000));
    expect(cellsOf(rowsOf(table)[0])[1].attrs.tcW).toEqual(dxa(1000));
  });
});

describe("new column width unit", () => {
  /**
   * The shape of a real contract: a table whose cells record their width as a percentage
   * of the grid total.
   * A percentage in `w:tcW` is in hundredths of a percent, so 5000 is 100%.
   */
  const pctCellsDoc = () =>
    tableDoc(
      [
        row(
          cell("Item", { tcW: pct(611) }),
          cell("Detail", { tcW: pct(3765) }),
          cell("Note", { tcW: pct(624) })
        ),
      ],
      { gridCols: [1110, 6840, 1134], tblW: dxa(9084) }
    );

  const dxaCellsDoc = () =>
    tableDoc(
      [
        row(
          cell("PartyA", { tcW: dxa(1110) }),
          cell("PartyB", { tcW: dxa(6840) })
        ),
      ],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );

  it("converts the dxa grid width to a percentage for pct cells", () => {
    const state = addColumnAfterCellWith(pctCellsDoc(), "Item");
    const { table } = firstTable(state.doc);
    const appended = cellsOf(rowsOf(table)[0])[1];

    // 1110 / 9084 * 5000 = 611. Using the grid width as is would put 1110 where a
    // percentage belongs.
    expect(appended.attrs.tcW).toEqual(pct(611));
  });

  it("keeps the grid width as is for dxa cells", () => {
    const state = addColumnAfterCellWith(dxaCellsDoc(), "PartyA");
    const { table } = firstTable(state.doc);
    const appended = cellsOf(rowsOf(table)[0])[1];

    // The unit stays dxa. The value is the reference column's 1110 after being shrunk by
    // the redistribution.
    expect(table.attrs.gridCols).toEqual([974, 974, 6002]);
    expect(appended.attrs.tcW).toEqual(dxa(974));
  });

  it("leaves percentage cell widths alone while the grid is reshared", () => {
    const state = addColumnAfterCellWith(pctCellsDoc(), "Item");
    const { table } = firstTable(state.doc);

    // A percentage is a ratio of the table width, so its value holds even as the grid
    // shrinks.
    expect(gridSum(table)).toBe(9084);
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([pct(611), pct(611), pct(3765), pct(624)]);
  });
});

describe("table width follows the grid total", () => {
  const dxaTableDoc = () =>
    tableDoc(
      [
        row(
          cell("A", { tcW: dxa(1110) }),
          cell("B", { tcW: dxa(6840) }),
          cell("C", { tcW: dxa(1134) })
        ),
      ],
      { gridCols: [1110, 6840, 1134], tblW: dxa(9084) }
    );

  it("keeps a dxa table width and reshares the grid instead", () => {
    const state = addColumnAfterCellWith(dxaTableDoc(), "A");
    const { table } = firstTable(state.doc);

    // Even with a column added, the table does not go past the page width.
    expect(table.attrs.gridCols).toEqual([989, 989, 6095, 1011]);
    expect(gridSum(table)).toBe(9084);
    expect(table.attrs.tblW).toEqual(dxa(9084));
  });

  it("shrinks a dxa table width by the deleted column", () => {
    const state = deleteColumnOfCellWith(dxaTableDoc(), "B");
    const { table } = firstTable(state.doc);

    expect(table.attrs.gridCols).toEqual([1110, 1134]);
    expect(table.attrs.tblW).toEqual(dxa(2244));
  });

  it("leaves a pct table alone (its width is a share of the page, not a sum)", () => {
    const doc = tableDoc(
      [row(cell("A", { tcW: pct(2500) }), cell("B", { tcW: pct(2500) }))],
      { gridCols: [100, 100], tblW: pct(5000) }
    );
    const state = addColumnAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);

    // Even as the grid grows, the table is still 100% of the page. There is no reason to
    // shrink it, so it just grows.
    expect(table.attrs.gridCols).toEqual([100, 100, 100]);
    expect(table.attrs.tblW).toEqual(pct(5000));
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([pct(2500), pct(2500), pct(2500)]);
  });
});

/**
 * The contract that prevents the old problem of a table spilling past the screen and the
 * page whenever a column was added.
 * The new column takes the reference column's width, but the whole grid is scaled down
 * proportionally so the grid total stays what it was before the insertion.
 */
describe("adding a column keeps the grid total", () => {
  const spannedDoc = () =>
    tableDoc(
      [
        row(
          cell("A", { tcW: dxa(1000) }),
          cell("B", { tcW: dxa(2000) }),
          cell("C", { tcW: dxa(3000) })
        ),
        row(
          cell("a", { tcW: dxa(1000) }),
          cell("bc", { colspan: 2, tcW: dxa(5000) })
        ),
      ],
      { gridCols: [1000, 2000, 3000], tblW: dxa(6000) }
    );

  const insertPoints: ReadonlyArray<[string, string]> = [
    ["before", "A"],
    ["before", "B"],
    ["before", "C"],
    ["after", "A"],
    ["after", "B"],
    ["after", "C"],
  ];

  it.each(insertPoints)("holds when inserting %s %s", (side, needle) => {
    const doc = spannedDoc();
    const state =
      side === "before"
        ? addColumnBeforeCellWith(doc, needle)
        : addColumnAfterCellWith(doc, needle);
    const { table } = firstTable(state.doc);

    expect(TableMap.get(table).width).toBe(4);
    expect(fixTables(state)).toBeUndefined();
    expect(gridSum(table)).toBe(6000);
    expect(table.attrs.tblW).toEqual(dxa(6000));
  });

  it("scales the cell widths with the grid and sums them for a merged cell", () => {
    // The new column goes right into the middle of the merged cell "bc" (grid columns
    // 1 to 2).
    const state = addColumnAfterCellWith(spannedDoc(), "B");
    const { table } = firstTable(state.doc);

    expect(table.attrs.gridCols).toEqual([750, 1500, 1500, 2250]);
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([dxa(750), dxa(1500), dxa(1500), dxa(2250)]);
    // The single cell covers three columns, so its width is the new sum of those three.
    expect(cellWithText(table, "bc")?.attrs).toMatchObject({
      colspan: 3,
      tcW: dxa(5250),
    });
  });

  it("puts the rounding leftover on the last column so the total is exact", () => {
    const doc = tableDoc(
      [
        row(
          cell("A", { tcW: dxa(333) }),
          cell("B", { tcW: dxa(667) }),
          cell("C", { tcW: dxa(1000) })
        ),
      ],
      { gridCols: [333, 667, 1000], tblW: dxa(2000) }
    );
    const state = addColumnAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);

    // Scaling by 2000/2333 leaves the last column at 857.26, but it takes on the 1 the
    // earlier columns dropped and ends up at 858.
    expect(table.attrs.gridCols).toEqual([285, 285, 572, 858]);
    expect(gridSum(table)).toBe(2000);
  });

  it("keeps a cell without a width empty and still reshares the grid", () => {
    const doc = tableDoc([row(cell("A"), cell("B", { tcW: dxa(2000) }))], {
      gridCols: [1000, 2000],
      tblW: dxa(3000),
    });
    const state = addColumnAfterCellWith(doc, "A");
    const { table } = firstTable(state.doc);

    // A cell with no width gets none created, because its original XML has to be written
    // back as it is.
    expect(table.attrs.gridCols).toEqual([750, 750, 1500]);
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([null, null, dxa(1500)]);
  });

  it("leaves a table without a grid alone", () => {
    const state = addColumnAfterCellWith(mergedContractDoc(), "Company");
    const { table } = firstTable(state.doc);

    // With no grid there is no width to redistribute either. The cell widths are exactly
    // what was inherited from the reference cell.
    expect(table.attrs.gridCols).toBeNull();
    expect(cellWithText(table, "Company")?.attrs.tcW).toEqual(dxa(1500));
    expect(cellsOf(rowsOf(table)[0])[2].attrs.tcW).toEqual(dxa(1500));
  });
});

describe("buildDeleteColumnTransaction", () => {
  it("refuses to delete the only column", () => {
    const doc = tableDoc([row(cell("Single", { tcW: dxa(1000) }))]);

    expect(
      buildDeleteColumnTransaction(stateWithCursorIn(doc, "Single"))
    ).toBeNull();
  });

  it("returns null outside a table", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("Body")),
    ]);

    expect(
      buildDeleteColumnTransaction(EditorState.create({ doc, schema }))
    ).toBeNull();
  });
});
