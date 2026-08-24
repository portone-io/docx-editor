// @vitest-environment node
import { describe, expect, it } from "vitest";
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
