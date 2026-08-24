// @vitest-environment node
import { describe, expect, it } from "vitest";
import { columnEdgeAt } from "./columnResize";

/** An ordinary cell covering grid column 1 on its own */
const single = { from: 1, to: 2 };
/** A merged cell covering grid columns 1 through 2 */
const merged = { from: 1, to: 3 };
const rect = { left: 100, right: 200 };

describe("columnEdgeAt", () => {
  it("treats a cell's right edge as the boundary of the last column it covers", () => {
    const right = { col: 1, side: "right" };
    expect(columnEdgeAt(200, rect, single, 4)).toEqual(right);
    expect(columnEdgeAt(202, rect, single, 4)).toEqual(right);
    expect(columnEdgeAt(197, rect, single, 4)).toEqual(right);
  });

  it("treats a cell's left edge as the boundary of the column before it", () => {
    const left = { col: 0, side: "left" };
    expect(columnEdgeAt(100, rect, single, 4)).toEqual(left);
    expect(columnEdgeAt(103, rect, single, 4)).toEqual(left);
  });

  it("points a merged cell at the outer boundaries of the grid span it covers", () => {
    expect(columnEdgeAt(200, rect, merged, 4)).toEqual({
      col: 2,
      side: "right",
    });
    expect(columnEdgeAt(100, rect, merged, 4)).toEqual({
      col: 0,
      side: "left",
    });
  });

  it("grabs nothing when the pointer is far from an edge", () => {
    expect(columnEdgeAt(150, rect, single, 4)).toBeNull();
    expect(columnEdgeAt(205, rect, single, 4)).toBeNull();
    expect(columnEdgeAt(95, rect, single, 4)).toBeNull();
  });

  it("does not grab the left end of the table", () => {
    expect(columnEdgeAt(100, rect, { from: 0, to: 1 }, 4)).toBeNull();
  });

  it("lets the right edge win when both edges are in reach in a narrow cell", () => {
    expect(columnEdgeAt(101, { left: 100, right: 104 }, single, 4)).toEqual({
      col: 1,
      side: "right",
    });
  });
});
