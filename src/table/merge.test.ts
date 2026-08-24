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
  schema,
  stateWithCellsSelected,
  stateWithCursorIn,
  tableDoc,
  uniformDoc,
} from "./__testing__/tables";
import { splitCell } from "./commands";
import {
  buildMergeCellsTransaction,
  buildSplitCellTransaction,
  canMergeCells,
  canSplitCell,
} from "./merge";

function mergeCellsFrom(
  doc: PMNode,
  anchorText: string,
  headText: string
): EditorState {
  const state = stateWithCellsSelected(doc, anchorText, headText);
  const tr = buildMergeCellsTransaction(state);
  if (!tr) throw new Error("merge refused");
  return state.apply(tr);
}

function splitCellOf(doc: PMNode, needle: string): EditorState {
  const state = stateWithCursorIn(doc, needle);
  const tr = buildSplitCellTransaction(state);
  if (!tr) throw new Error("split refused");
  return state.apply(tr);
}

/** Extracts just the table's shape. Used for round-trip comparison */
function gridShape(table: PMNode): {
  size: [number, number];
  spans: [number, number][];
} {
  const map = TableMap.get(table);
  const spans: [number, number][] = [];
  table.descendants((node) => {
    if (node.type.spec.tableRole !== "cell") return true;
    spans.push([node.attrs.colspan, node.attrs.rowspan]);
    return false;
  });
  return { size: [map.width, map.height], spans };
}

describe("canMergeCells", () => {
  it("accepts a rectangular selection of two neighbours", () => {
    expect(
      canMergeCells(stateWithCellsSelected(uniformDoc(), "Item", "Detail"))
    ).toBe(true);
  });

  it("rejects a plain cursor", () => {
    expect(canMergeCells(stateWithCursorIn(uniformDoc(), "Item"))).toBe(false);
  });

  it("rejects a single selected cell", () => {
    expect(
      canMergeCells(stateWithCellsSelected(uniformDoc(), "Item", "Item"))
    ).toBe(false);
  });

  it("rejects a selection a merged cell sticks out of", () => {
    // From Company to 555-0123 covers grid columns 1 to 2.
    // A colspan-2 cell starting inside that range reaches out to column 3 and sticks out
    // of the selection.
    expect(
      canMergeCells(
        stateWithCellsSelected(mergedContractDoc(), "Company", "555-0123")
      )
    ).toBe(false);
  });

  it("builds nothing when it cannot merge", () => {
    expect(
      buildMergeCellsTransaction(stateWithCursorIn(uniformDoc(), "Item"))
    ).toBeNull();
  });
});

describe("buildMergeCellsTransaction", () => {
  it("merges a row of cells into one and keeps the grid valid", () => {
    const state = mergeCellsFrom(uniformDoc(), "Item", "Amount");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(3);
    expect(map.height).toBe(3);
    expect(fixTables(state)).toBeUndefined();

    const merged = cellsOf(rowsOf(table)[0]);
    expect(merged.length).toBe(1);
    expect(merged[0].attrs.colspan).toBe(3);
    expect(merged[0].attrs.rowspan).toBe(1);
    expect(merged[0].textContent).toBe("ItemDetailAmount");
  });

  it("merges a 2x2 block", () => {
    const state = mergeCellsFrom(uniformDoc(), "Design", "Implementation");
    const { table } = firstTable(state.doc);

    expect(fixTables(state)).toBeUndefined();
    expect(TableMap.get(table).height).toBe(3);
    expect(
      cellWithText(table, "DesignRequirementsBuildImplementation")?.attrs
    ).toMatchObject({ colspan: 2, rowspan: 2 });
  });

  it("keeps the top-left cell's format and drops the others'", () => {
    const doc = tableDoc([
      row(
        cell("Left", {
          tcAttrs: ' w:id="1"',
          tcPr: '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>',
          tcW: dxa(1000),
          format: { background: "#ffff00" },
        }),
        cell("Right", {
          tcAttrs: ' w:id="2"',
          tcPr: '<w:tcPr><w:shd w:fill="00FF00"/></w:tcPr>',
          tcW: dxa(2000),
          format: { background: "#00ff00" },
        })
      ),
    ]);
    const state = mergeCellsFrom(doc, "Left", "Right");
    const { table } = firstTable(state.doc);
    const merged = cellsOf(rowsOf(table)[0])[0];

    expect(merged.attrs.tcAttrs).toBe(' w:id="1"');
    expect(merged.attrs.tcPr).toBe('<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>');
    expect(merged.attrs.format).toEqual({ background: "#ffff00" });
    expect(merged.attrs.colspan).toBe(2);
  });

  it("widens the merged cell to the grid columns it now covers", () => {
    const doc = tableDoc(
      [row(cell("A", { tcW: dxa(1110) }), cell("B", { tcW: dxa(6840) }))],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );
    const state = mergeCellsFrom(doc, "A", "B");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    expect(merged.attrs.tcW).toEqual(dxa(7950));
  });

  it("converts the covered grid width into the merged cell's own unit", () => {
    const doc = tableDoc(
      [
        row(
          cell("A", { tcW: pct(698) }),
          cell("B", { tcW: pct(2151) }),
          cell("C", { tcW: pct(2151) })
        ),
      ],
      { gridCols: [1110, 3420, 3420], tblW: dxa(7950) }
    );
    const state = mergeCellsFrom(doc, "A", "B");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    // (1110 + 3420) / 7950 * 5000 = 2849
    expect(merged.attrs.tcW).toEqual(pct(2849));
  });

  it("adds up the covered cells' widths when the table has no grid", () => {
    const doc = tableDoc([
      row(
        cell("A", { tcW: dxa(1000) }),
        cell("B", { tcW: dxa(2000) }),
        cell("C", { tcW: dxa(3000) })
      ),
    ]);
    const state = mergeCellsFrom(doc, "A", "B");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    expect(merged.attrs.tcW).toEqual(dxa(3000));
  });

  it("counts a spanning cell once when adding up covered widths", () => {
    const doc = tableDoc([
      row(
        cell("Merged", { colspan: 2, tcW: dxa(3000) }),
        cell("End", { tcW: dxa(1000) })
      ),
      row(cell("Below", { colspan: 3, tcW: dxa(4000) })),
    ]);
    const state = mergeCellsFrom(doc, "Merged", "End");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    expect(merged.attrs.tcW).toEqual(dxa(4000));
  });

  it("keeps the column width when the merge only spans rows", () => {
    const doc = tableDoc(
      [
        row(
          cell("Above", { tcW: dxa(1110) }),
          cell("Side", { tcW: dxa(6840) })
        ),
        row(
          cell("Below", { tcW: dxa(1110) }),
          cell("Side2", { tcW: dxa(6840) })
        ),
      ],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );
    const state = mergeCellsFrom(doc, "Above", "Below");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    expect(merged.attrs).toMatchObject({ rowspan: 2, tcW: dxa(1110) });
  });

  it("leaves a cell without a width without one", () => {
    const doc = tableDoc([row(cell("A"), cell("B", { tcW: dxa(2000) }))], {
      gridCols: [1000, 2000],
    });
    const state = mergeCellsFrom(doc, "A", "B");
    const merged = cellsOf(rowsOf(firstTable(state.doc).table)[0])[0];

    expect(merged.attrs.tcW).toBeNull();
  });

  it("leaves the grid columns and the table width alone", () => {
    const doc = tableDoc(
      [row(cell("A", { tcW: dxa(1000) }), cell("B", { tcW: dxa(2000) }))],
      { gridCols: [1000, 2000], tblW: dxa(3000) }
    );
    const state = mergeCellsFrom(doc, "A", "B");
    const { table } = firstTable(state.doc);

    expect(table.attrs.gridCols).toEqual([1000, 2000]);
    expect(table.attrs.tblW).toEqual(dxa(3000));
  });
});

describe("canSplitCell", () => {
  it("accepts a merged cell", () => {
    expect(canSplitCell(stateWithCursorIn(mergedContractDoc(), "PartyB"))).toBe(
      true
    );
  });

  it("rejects a plain 1x1 cell", () => {
    expect(canSplitCell(stateWithCursorIn(uniformDoc(), "Item"))).toBe(false);
  });

  it("rejects a cursor outside a table", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("Body")),
    ]);
    const state = EditorState.create({ doc, schema });

    expect(canSplitCell(state)).toBe(false);
    expect(buildSplitCellTransaction(state)).toBeNull();
  });
});

/**
 * Upstream builds the cells of a split out of the original cell's attrs, so each of them would
 * carry the original's lock: a lock planted where it was never put, which the guard refuses
 * (`schema/locks`). `writeSplitFormats` takes the lock back off every spot but the original's, and
 * the refusal stands all the same, because the step that plants it comes first. So a merged cell a
 * control shuts stays whole until the lock is lifted, and the query has to say so.
 */
describe("splitting a merged cell a control shuts", () => {
  const lockedMergedDoc = () =>
    tableDoc(
      [
        row(
          cell("Shut", {
            colspan: 2,
            tcW: dxa(2000),
            sdtPrefix: "<w:sdt><w:sdtPr/>",
            sdtContentsLocked: true,
          })
        ),
        row(cell("A", { tcW: dxa(1000) }), cell("B", { tcW: dxa(1000) })),
      ],
      { gridCols: [1000, 1000] }
    );

  it("is refused, though the grid allows it", () => {
    const state = stateWithCursorIn(lockedMergedDoc(), "Shut");
    // The grid has a merged cell to split: it is the lock that turns the split down
    expect(buildSplitCellTransaction(state)).not.toBeNull();
    expect(canSplitCell(state)).toBe(false);
    expect(splitCell(state)).toBe(false);
  });

  it("is offered again on the same cell once no control shuts it", () => {
    const doc = tableDoc(
      [
        row(cell("Open", { colspan: 2, tcW: dxa(2000) })),
        row(cell("A", { tcW: dxa(1000) }), cell("B", { tcW: dxa(1000) })),
      ],
      { gridCols: [1000, 1000] }
    );
    const state = stateWithCursorIn(doc, "Open");

    expect(canSplitCell(state)).toBe(true);
    expect(splitCell(state)).toBe(true);
  });
});

describe("buildSplitCellTransaction", () => {
  it("turns a colspan cell into plain cells and keeps the grid valid", () => {
    const state = splitCellOf(mergedContractDoc(), "Acme Corp");
    const { table } = firstTable(state.doc);
    const map = TableMap.get(table);

    expect(map.width).toBe(4);
    expect(map.height).toBe(3);
    expect(fixTables(state)).toBeUndefined();
    // The two cells produced by the split stand alone, and the neighbouring cells
    // (vertical merges included) are unchanged.
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => [
        current.attrs.colspan,
        current.attrs.rowspan,
      ])
    ).toEqual([
      [1, 3],
      [1, 1],
      [1, 1],
      [1, 1],
    ]);
  });

  it("turns a rowspan cell into one cell per row", () => {
    const state = splitCellOf(mergedContractDoc(), "PartyB");
    const { table } = firstTable(state.doc);

    expect(fixTables(state)).toBeUndefined();
    expect(rowsOf(table).map((current) => current.childCount)).toEqual([
      3, 3, 4,
    ]);
    expect(cellWithText(table, "PartyB")?.attrs.rowspan).toBe(1);
  });

  it("gives each new cell the format of the split cell", () => {
    const doc = tableDoc([
      row(
        cell("Merged", {
          colspan: 2,
          tcAttrs: ' w:id="1"',
          tcPr: '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>',
          tcW: dxa(3000),
          format: { background: "#ffff00" },
        })
      ),
      row(cell("Left", { tcW: dxa(1000) }), cell("Right", { tcW: dxa(2000) })),
    ]);
    const state = splitCellOf(doc, "Merged");
    const { table } = firstTable(state.doc);
    const split = cellsOf(rowsOf(table)[0]);

    expect(split.length).toBe(2);
    for (const current of split) {
      expect(current.attrs.tcAttrs).toBe(' w:id="1"');
      expect(current.attrs.tcPr).toBe(
        '<w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>'
      );
      expect(current.attrs.format).toEqual({ background: "#ffff00" });
      expect(current.attrs.colspan).toBe(1);
      expect(current.attrs.rowspan).toBe(1);
    }
  });

  it("leaves the content control on the split cell alone and copies it nowhere", () => {
    const doc = tableDoc([
      row(
        cell("Merged", {
          colspan: 2,
          sdtPrefix: "<w:sdt><w:sdtPr/>",
          sdtContentsLocked: true,
        })
      ),
      row(cell("Left"), cell("Right")),
    ]);
    const state = splitCellOf(doc, "Merged");
    const cells = cellsOf(rowsOf(firstTable(state.doc).table)[0]);

    expect(cells.map((current) => current.attrs.sdtPrefix)).toEqual([
      "<w:sdt><w:sdtPr/>",
      null,
    ]);
    // The lock lives inside that control, so a cell without one is not locked either
    expect(cells.map((current) => current.attrs.sdtContentsLocked)).toEqual([
      true,
      false,
    ]);
  });

  it("takes each new cell's width from its own grid column", () => {
    const doc = tableDoc(
      [
        row(cell("Merged", { colspan: 2, tcW: dxa(9000) })),
        row(
          cell("Left", { tcW: dxa(1110) }),
          cell("Right", { tcW: dxa(6840) })
        ),
      ],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );
    const state = splitCellOf(doc, "Merged");
    const { table } = firstTable(state.doc);

    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([dxa(1110), dxa(6840)]);
    // A split never changes the grid.
    expect(table.attrs.gridCols).toEqual([1110, 6840]);
    expect(table.attrs.tblW).toEqual(dxa(7950));
  });

  it("converts the grid width into the split cell's own unit", () => {
    const doc = tableDoc(
      [
        row(cell("Merged", { colspan: 2, tcW: pct(5000) })),
        row(cell("Left", { tcW: pct(611) }), cell("Right", { tcW: pct(4389) })),
      ],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );
    const state = splitCellOf(doc, "Merged");
    const { table } = firstTable(state.doc);

    // 1110 / 7950 * 5000 = 698, 6840 / 7950 * 5000 = 4302
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([pct(698), pct(4302)]);
  });

  it("splits the reference width by the span when the table has no grid", () => {
    const doc = tableDoc([
      row(cell("Merged", { colspan: 2, tcW: dxa(4000) })),
      row(cell("Left", { tcW: dxa(2000) }), cell("Right", { tcW: dxa(2000) })),
    ]);
    const state = splitCellOf(doc, "Merged");
    const { table } = firstTable(state.doc);

    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([dxa(2000), dxa(2000)]);
  });
});

describe("merge and split round trip", () => {
  it("returns the grid to its original shape", () => {
    const before = uniformDoc();
    const merged = mergeCellsFrom(before, "Design", "Implementation");
    const restored = splitCellOf(merged.doc, "Design");

    expect(fixTables(restored)).toBeUndefined();
    expect(gridShape(firstTable(restored.doc).table)).toEqual(
      gridShape(firstTable(before).table)
    );
  });

  it("keeps the widths of a merged then split column", () => {
    const before = tableDoc(
      [
        row(cell("A", { tcW: dxa(1110) }), cell("B", { tcW: dxa(6840) })),
        row(cell("a", { tcW: dxa(1110) }), cell("b", { tcW: dxa(6840) })),
      ],
      { gridCols: [1110, 6840], tblW: dxa(7950) }
    );
    const merged = mergeCellsFrom(before, "A", "B");
    const restored = splitCellOf(merged.doc, "A");
    const { table } = firstTable(restored.doc);

    expect(fixTables(restored)).toBeUndefined();
    expect(
      cellsOf(rowsOf(table)[0]).map((current) => current.attrs.tcW)
    ).toEqual([dxa(1110), dxa(6840)]);
    expect(table.attrs.gridCols).toEqual([1110, 6840]);
    expect(table.attrs.tblW).toEqual(dxa(7950));
  });
});
