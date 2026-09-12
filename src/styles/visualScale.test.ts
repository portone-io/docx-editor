// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { editorClassNames } from "./classNames";
import { pageScaleAround, visualScaleOf } from "./visualScale";

function element(transform?: string): HTMLElement {
  const made = document.createElement("div");
  if (transform !== undefined) made.style.transform = transform;
  document.body.append(made);
  return made;
}

describe("visualScaleOf", () => {
  it("reads the horizontal scale out of the matrix a browser computes", () => {
    expect(visualScaleOf(element("matrix(0.75, 0, 0, 0.75, 0, 0)"))).toBe(0.75);
    expect(
      visualScaleOf(
        element("matrix3d(0.5, 0, 0, 0, 0, 0.5, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1)")
      )
    ).toBe(0.5);
  });

  it("reads it out of the function itself where none was computed", () => {
    expect(visualScaleOf(element("scale(0.6)"))).toBe(0.6);
    expect(visualScaleOf(element("scale(1.25, 1.25)"))).toBe(1.25);
    expect(visualScaleOf(element("scaleX(2)"))).toBe(2);
  });

  it("answers 1 for anything that scales the element by nothing it can read", () => {
    expect(visualScaleOf(element())).toBe(1);
    expect(visualScaleOf(element("none"))).toBe(1);
    expect(visualScaleOf(element("translate(40px, 0)"))).toBe(1);
    // A scale of zero would divide a measurement by nothing
    expect(visualScaleOf(element("scale(0)"))).toBe(1);
    expect(visualScaleOf(null)).toBe(1);
  });
});

describe("pageScaleAround", () => {
  it("takes the scale from the page layer a node is drawn inside", () => {
    const layer = element("scale(0.5)");
    layer.className = editorClassNames.pageLayer;
    const sheet = document.createElement("div");
    layer.append(sheet);

    expect(pageScaleAround(sheet)).toBe(0.5);
  });

  it("answers 1 for a node drawn outside any page layer", () => {
    const scaled = element("scale(0.5)");
    const inside = document.createElement("div");
    scaled.append(inside);

    expect(pageScaleAround(inside)).toBe(1);
  });
});
