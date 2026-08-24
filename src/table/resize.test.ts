// @vitest-environment node
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { fixTables, TableMap } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { LETTER_GEOMETRY as LETTER } from "../__testing__/docx";
import {
  A4_BODY_WIDTH,
  A4_PORTRAIT,
  type PageGeometry,
} from "../docx/pageGeometry";
import {
  cell,
  cellsOf,
  cellWithText,
  dxa,
  firstTable,
  pct,
  row,
  rowsOf,
  schema,
  tableDoc,
} from "./__testing__/tables";
import {
  buildResizeColumnTransaction,
  MIN_COLUMN_DXA,
  maxGridTotal,
  resizedGridCols,
} from "./resize";

/** The sum of the grid column widths */
function gridSum(table: PMNode): number {
  const value: unknown = table.attrs.gridCols;
  const cols = Array.isArray(value) ? value : [];
  return cols.reduce<number>(
    (sum, width) => sum + (typeof width === "number" ? width : 0),
    0
  );
}

function resizeFirstTable(
  doc: PMNode,
  col: number,
  width: number,
  geometry: PageGeometry = A4_PORTRAIT
): EditorState | null {
  const state = EditorState.create({ doc, schema });
  const { start } = firstTable(doc);
  const tr = buildResizeColumnTransaction(
    state,
    { tablePos: start - 1, col, width },
    geometry
  );
  return tr ? state.apply(tr) : null;
}

function resizedTable(
  doc: PMNode,
  col: number,
  width: number,
  geometry?: PageGeometry
): PMNode {
  const state = resizeFirstTable(doc, col, width, geometry);
  if (!state) throw new Error("could not move the boundary");
  return firstTable(state.doc).table;
}

/** A three-column table whose grid total equals its dxa table width */
const evenDoc = () =>
  tableDoc(
    [
      row(
        cell("A", { tcW: dxa(2000) }),
        cell("B", { tcW: dxa(3000) }),
        cell("C", { tcW: dxa(1000) })
      ),
    ],
    { gridCols: [2000, 3000, 1000], tblW: dxa(6000) }
  );

describe("resizedGridCols - inner boundary", () => {
  it("only the two touching columns trade width and the total stays the same", () => {
    expect(
      resizedGridCols([2000, 3000, 1000], 0, 2500, A4_BODY_WIDTH.twips)
    ).toEqual([2500, 2500, 1000]);
  });

  it("dragging to the left widens the right-hand column by that much", () => {
    expect(
      resizedGridCols([2000, 3000, 1000], 1, 2400, A4_BODY_WIDTH.twips)
    ).toEqual([2000, 2400, 1600]);
  });

  it("is null when nothing changes", () => {
    expect(
      resizedGridCols([2000, 3000, 1000], 0, 2000, A4_BODY_WIDTH.twips)
    ).toBeNull();
  });

  it("a boundary out of range is null", () => {
    expect(
      resizedGridCols([2000, 1000], -1, 500, A4_BODY_WIDTH.twips)
    ).toBeNull();
    expect(
      resizedGridCols([2000, 1000], 2, 500, A4_BODY_WIDTH.twips)
    ).toBeNull();
  });
});

describe("resizedGridCols - clamping", () => {
  it("cannot narrow its own column below the minimum width", () => {
    const resized = resizedGridCols([2000, 3000], 0, 10, A4_BODY_WIDTH.twips);
    expect(resized).toEqual([MIN_COLUMN_DXA, 5000 - MIN_COLUMN_DXA]);
  });

  it("cannot narrow the neighbouring column below the minimum width", () => {
    const resized = resizedGridCols([2000, 3000], 0, 9000, A4_BODY_WIDTH.twips);
    expect(resized).toEqual([5000 - MIN_COLUMN_DXA, MIN_COLUMN_DXA]);
  });

  it("for a table whose grid is written down only as ratios, the floor drops to match those ratios too", () => {
    // The total is only 300, so keeping the 240 dxa floor as is would leave every
    // boundary immovable. The floor is 50, half of an evenly divided width (100).
    expect(
      resizedGridCols([100, 100, 100], 0, 150, A4_BODY_WIDTH.twips)
    ).toEqual([150, 50, 100]);
    expect(
      resizedGridCols([100, 100, 100], 0, 999, A4_BODY_WIDTH.twips)
    ).toEqual([150, 50, 100]);
  });

  it("does not move the end boundary when the room left is narrower than the floor", () => {
    expect(resizedGridCols([9000, 300], 1, 500, 9100)).toBeNull();
  });

  it("the end boundary cannot go past the body width", () => {
    const resized = resizedGridCols(
      [2000, 3000],
      1,
      90_000,
      A4_BODY_WIDTH.twips
    );
    expect(resized).toEqual([2000, A4_BODY_WIDTH.twips - 2000]);
  });

  it("the end boundary does not narrow below the minimum width either", () => {
    expect(resizedGridCols([2000, 3000], 1, 0, A4_BODY_WIDTH.twips)).toEqual([
      2000,
      MIN_COLUMN_DXA,
    ]);
  });
});

describe("maxGridTotal", () => {
  it("the body width is the limit when there is no indent", () => {
    const { table } = firstTable(evenDoc());
    expect(maxGridTotal(table, A4_PORTRAIT)).toBe(A4_BODY_WIDTH.twips);
  });

  it("the limit shrinks by the table indent", () => {
    // indentLeftPt 50pt = 1000 dxa
    const doc = tableDoc([row(cell("A", { tcW: dxa(1000) }))], {
      gridCols: [1000],
      tblW: dxa(1000),
      format: { indentLeftPt: 50 },
    });
    expect(maxGridTotal(firstTable(doc).table, A4_PORTRAIT)).toBe(
      A4_BODY_WIDTH.twips - 1000
    );
  });

  it("for a table already wider than the body, its current width is the limit", () => {
    const wide = A4_BODY_WIDTH.twips + 2000;
    const doc = tableDoc([row(cell("A", { tcW: dxa(wide) }))], {
      gridCols: [wide],
      tblW: dxa(wide),
    });
    expect(maxGridTotal(firstTable(doc).table, A4_PORTRAIT)).toBe(wide);
  });

  it("the limit is the paper's own body width, not A4's", () => {
    const { table } = firstTable(evenDoc());
    // Letter with an inch of margin leaves 9360 dxa, A4 with 2.2cm leaves 9412
    expect(maxGridTotal(table, LETTER)).toBe(9360);
    expect(maxGridTotal(table, LETTER)).not.toBe(
      maxGridTotal(table, A4_PORTRAIT)
    );
  });

  it("a column dragged past the end stops at the paper's own body width", () => {
    const table = resizedTable(evenDoc(), 2, 90_000, LETTER);
    expect(gridSum(table)).toBe(9360);
    expect(table.attrs.tblW).toEqual(dxa(9360));
  });
});

describe("buildResizeColumnTransaction - inner boundary", () => {
  it("the grid total and the table width stay the same", () => {
    const table = resizedTable(evenDoc(), 0, 2500);

    expect(table.attrs.gridCols).toEqual([2500, 2500, 1000]);
    expect(gridSum(table)).toBe(6000);
    expect(table.attrs.tblW).toEqual(dxa(6000));
  });

  it("only the widths of the two touching cells change", () => {
    const table = resizedTable(evenDoc(), 0, 2500);

    expect(cellsOf(rowsOf(table)[0]).map((one) => one.attrs.tcW)).toEqual([
      dxa(2500),
      dxa(2500),
      dxa(1000),
    ]);
  });

  it("finishes in a single step (one undo puts it back)", () => {
    const doc = evenDoc();
    const state = EditorState.create({ doc, schema });
    const tr = buildResizeColumnTransaction(
      state,
      { tablePos: firstTable(doc).start - 1, col: 0, width: 2500 },
      A4_PORTRAIT
    );

    expect(tr).not.toBeNull();
    expect(tr?.docChanged).toBe(true);
  });

  it("leaves a table whose grid is unknown untouched", () => {
    const doc = tableDoc([row(cell("A", { tcW: dxa(1000) }))], {
      gridCols: [],
      tblW: dxa(1000),
    });

    expect(resizeFirstTable(doc, 0, 2000)).toBeNull();
  });

  it("a position that is not a table is null", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("Body")),
    ]);
    const state = EditorState.create({ doc, schema });

    expect(
      buildResizeColumnTransaction(
        state,
        { tablePos: 0, col: 0, width: 100 },
        A4_PORTRAIT
      )
    ).toBeNull();
  });
});

describe("buildResizeColumnTransaction - end boundary", () => {
  it("the total width changes and a dxa table width follows it", () => {
    const table = resizedTable(evenDoc(), 2, 2000);

    expect(table.attrs.gridCols).toEqual([2000, 3000, 2000]);
    expect(gridSum(table)).toBe(7000);
    expect(table.attrs.tblW).toEqual(dxa(7000));
  });

  it("narrowing it narrows the table along with it", () => {
    const table = resizedTable(evenDoc(), 2, 500);

    expect(gridSum(table)).toBe(5500);
    expect(table.attrs.tblW).toEqual(dxa(5500));
  });

  it("leaves a percentage table width alone and changes only the grid ratios", () => {
    const doc = tableDoc(
      [row(cell("A", { tcW: pct(2500) }), cell("B", { tcW: pct(2500) }))],
      { gridCols: [3000, 3000], tblW: pct(5000) }
    );
    const table = resizedTable(doc, 1, 4000);

    expect(table.attrs.gridCols).toEqual([3000, 4000]);
    expect(table.attrs.tblW).toEqual(pct(5000));
  });

  it("stops at the body width", () => {
    const table = resizedTable(evenDoc(), 2, 90_000);

    expect(gridSum(table)).toBe(A4_BODY_WIDTH.twips);
    expect(table.attrs.tblW).toEqual(dxa(A4_BODY_WIDTH.twips));
  });
});

describe("buildResizeColumnTransaction - cell widths", () => {
  const spannedDoc = () =>
    tableDoc(
      [
        row(
          cell("A", { tcW: dxa(2000) }),
          cell("B", { tcW: dxa(3000) }),
          cell("C", { tcW: dxa(1000) })
        ),
        row(
          cell("ab", { colspan: 2, tcW: dxa(5000) }),
          cell("c", { tcW: dxa(1000) })
        ),
      ],
      { gridCols: [2000, 3000, 1000], tblW: dxa(6000) }
    );

  it("a merged cell's width is the sum of the grid columns it covers", () => {
    // An inner boundary, so the two columns' total is unchanged and so is the merged cell
    const inner = resizedTable(spannedDoc(), 0, 2500);
    expect(cellWithText(inner, "ab")?.attrs.tcW).toEqual(dxa(5000));

    // Dragging the end boundary moves only a column the merged cell does not cover, so
    // the merged cell stays the same
    const end = resizedTable(spannedDoc(), 2, 2000);
    expect(cellWithText(end, "ab")?.attrs.tcW).toEqual(dxa(5000));
    expect(cellWithText(end, "c")?.attrs.tcW).toEqual(dxa(2000));
  });

  it("when a column a merged cell covers moves, it is written again as that sum", () => {
    const table = resizedTable(spannedDoc(), 1, 2000);

    // Grid column 1 shrank from 3000 to 2000 and column 2 grew by that much
    expect(table.attrs.gridCols).toEqual([2000, 2000, 2000]);
    expect(cellWithText(table, "ab")?.attrs.tcW).toEqual(dxa(4000));
  });

  it("leaves percentage cells and cells with no width alone", () => {
    const doc = tableDoc(
      [
        row(
          cell("A", { tcW: pct(2500) }),
          cell("B"),
          cell("C", { tcW: dxa(1000) })
        ),
      ],
      { gridCols: [2000, 3000, 1000], tblW: dxa(6000) }
    );
    const table = resizedTable(doc, 0, 2500);

    expect(cellsOf(rowsOf(table)[0]).map((one) => one.attrs.tcW)).toEqual([
      pct(2500),
      null,
      dxa(1000),
    ]);
  });

  it("the resized table has nothing to fix (fixTables finds no repair)", () => {
    const state = resizeFirstTable(spannedDoc(), 0, 2500);
    if (!state) throw new Error("could not move the boundary");

    expect(fixTables(state)).toBeUndefined();
    expect(TableMap.get(firstTable(state.doc).table).width).toBe(3);
  });
});
