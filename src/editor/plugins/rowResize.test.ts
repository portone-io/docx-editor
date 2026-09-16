// @vitest-environment node
import { describe, expect, it } from "vitest";
import {
  CELL_CONTROL_PREFIX,
  cell,
  firstTable,
  ROW_CONTROL_PREFIX,
  row,
  rowWith,
  tableDoc,
} from "../../table/__testing__/tables";
import { isRowResizable } from "../../table/rowResize";
import { rowEdgeAt } from "./rowResize";

const single = { from: 1, to: 2 };
const merged = { from: 1, to: 3 };
const rect = { top: 100, bottom: 200 };

describe("rowEdgeAt", () => {
  it("targets the row above a cell's bottom edge", () => {
    expect(rowEdgeAt(200, rect, single, 4)).toEqual({
      row: 1,
      side: "bottom",
    });
    expect(rowEdgeAt(197, rect, single, 4)).toEqual({
      row: 1,
      side: "bottom",
    });
  });

  it("targets the row before a cell's top edge", () => {
    expect(rowEdgeAt(102, rect, single, 4)).toEqual({
      row: 0,
      side: "top",
    });
  });

  it("uses the last covered row at a merged cell's bottom edge", () => {
    expect(rowEdgeAt(200, rect, merged, 4)).toEqual({
      row: 2,
      side: "bottom",
    });
  });

  it("does not grab the table's top edge or a distant pointer", () => {
    expect(rowEdgeAt(100, rect, { from: 0, to: 1 }, 4)).toBeNull();
    expect(rowEdgeAt(150, rect, single, 4)).toBeNull();
  });
});

/**
 * The affordance `edgeUnder` offers is decided by the same predicate the drag is, so a row the
 * drag would refuse never shows a handle.
 */
describe("the rows the pointer offers a handle on", () => {
  const tableOf = (doc: ReturnType<typeof tableDoc>) => firstTable(doc).table;

  it("offers one on an ordinary row", () => {
    expect(isRowResizable(tableOf(tableDoc([row(cell("A"))])), 0)).toBe(true);
  });

  it("offers none on a row a content control shuts", () => {
    const doc = tableDoc([
      rowWith(
        { sdtPrefix: ROW_CONTROL_PREFIX, sdtContentsLocked: true },
        cell("Shut")
      ),
    ]);
    expect(isRowResizable(tableOf(doc), 0)).toBe(false);
  });

  it("offers none on a row holding a cell a content control shuts", () => {
    const doc = tableDoc([
      row(
        cell("Locked", {
          sdtPrefix: CELL_CONTROL_PREFIX,
          sdtContentsLocked: true,
        })
      ),
    ]);
    expect(isRowResizable(tableOf(doc), 0)).toBe(false);
  });

  it("offers none beyond the rows the table holds", () => {
    const table = tableOf(tableDoc([row(cell("A"))]));
    expect(isRowResizable(table, -1)).toBe(false);
    expect(isRowResizable(table, 1)).toBe(false);
  });
});
