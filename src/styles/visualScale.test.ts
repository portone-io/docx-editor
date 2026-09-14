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

/** An element inside the one given, which carries a transform of its own or none */
function inside(parent: HTMLElement, transform?: string): HTMLElement {
  const made = document.createElement("div");
  if (transform !== undefined) made.style.transform = transform;
  parent.append(made);
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

  // An application is free to scale whatever it mounts the editor inside, and a rectangle read
  // within one of those is drawn at every scale over it multiplied together
  it("multiplies its own scale by every scale standing above it", () => {
    const shell = element("scale(0.5)");
    expect(visualScaleOf(inside(shell, "scale(0.6)"))).toBeCloseTo(0.3, 10);
    expect(visualScaleOf(inside(inside(shell), "scale(2)"))).toBe(1);
    expect(visualScaleOf(inside(shell))).toBe(0.5);
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

    expect(pageScaleAround(inside(layer))).toBe(0.5);
  });

  it("takes in what scales the editor as a whole as well", () => {
    const shell = element("scale(0.5)");
    const layer = inside(shell, "scale(0.6)");
    layer.className = editorClassNames.pageLayer;

    expect(pageScaleAround(inside(layer))).toBeCloseTo(0.3, 10);
  });

  it("answers 1 for a node drawn outside any page layer", () => {
    expect(pageScaleAround(inside(element("scale(0.5)")))).toBe(1);
  });
});
