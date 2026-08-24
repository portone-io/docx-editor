// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { editorAttributes } from "../styles/classNames";
import { measureSheet } from "./measureBlocks";
import { setPageBreakSpaces, setPagePushes } from "./pageDecorations";
import { type MeasuredBlock, pageLayout } from "./pageLayout";

const PAGE = 500;
const STEP = 100;
const WIDTH = 400;

let view: EditorView | null = null;
let layer: HTMLElement | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  layer?.remove();
  layer = null;
});

/** One block as it would be drawn with none of the engine's marks on it */
interface BlockShape {
  gap: number;
  height: number;
  /** Where the block's page breaks stand, measured from its own top */
  breaks: readonly number[];
}

function editor(doc: PMNode): EditorView {
  const mount = document.createElement("div");
  layer = mount;
  document.body.appendChild(mount);
  view = new EditorView(mount, { state: createEditorState(doc) });
  return view;
}

function pageBreak(): PMNode {
  return docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' });
}

function number(value: string | null): number {
  const parsed = Number.parseFloat(value ?? "");
  return Number.isFinite(parsed) ? parsed : 0;
}

function spaceElements(dom: HTMLElement): Element[] {
  return Array.from(
    dom.querySelectorAll(`[${editorAttributes.pageBreakSpace}]`)
  );
}

function rect(element: Element, top: number, height: number): void {
  element.getBoundingClientRect = () => new DOMRect(0, top, WIDTH, height);
}

/**
 * Draws the sheet as the browser would, from the shapes the blocks have of their own and the
 * marks the engine has put on them: a push widens the gap above a block, and a break's space
 * grows the block it sits in and moves everything after the break down inside it.
 */
function draw(
  live: EditorView,
  shapes: readonly BlockShape[],
  scale = 1
): void {
  let y = 0;
  live.state.doc.forEach((_node, offset, index) => {
    const shape = shapes[index];
    const dom = live.nodeDOM(offset);
    if (!shape || !(dom instanceof HTMLElement)) return;
    const top =
      y + shape.gap + number(dom.getAttribute(editorAttributes.pagePush));
    let opened = 0;
    spaceElements(dom).forEach((element, at) => {
      const height = number(
        element.getAttribute(editorAttributes.pageBreakSpace)
      );
      rect(
        element,
        (top + (shape.breaks[at] ?? 0) + opened) * scale,
        height * scale
      );
      opened += height;
    });
    rect(dom, top * scale, (shape.height + opened) * scale);
    y = top + shape.height + opened;
  });
}

function layoutOf(blocks: readonly MeasuredBlock[]) {
  return pageLayout({ blocks, pageBodyHeight: PAGE, pageStep: STEP });
}

/** A paragraph carrying a page break, and one below it tall enough to be pushed off the page */
function brokenParagraph(): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create({}, [
      docxSchema.text("before"),
      pageBreak(),
      docxSchema.text("after"),
    ]),
    docxSchema.nodes.paragraph.create({}, [docxSchema.text("next")]),
  ]);
}

const SHAPES: readonly BlockShape[] = [
  { gap: 0, height: 40, breaks: [20] },
  { gap: 0, height: PAGE, breaks: [] },
];

describe("measureSheet", () => {
  it("gives every page break a space to be measured at, before any measurement", () => {
    const live = editor(brokenParagraph());
    expect(spaceElements(live.dom)).toHaveLength(1);
    expect(
      spaceElements(live.dom)[0]?.getAttribute(editorAttributes.pageBreakSpace)
    ).toBe("0");
  });

  it("reads a break where it stands in its own block", () => {
    const live = editor(brokenParagraph());
    draw(live, SHAPES);
    expect(measureSheet(live, live.dom).blocks).toEqual([
      { pos: 0, gap: 0, height: 40, breakBefore: false, breaks: [20] },
      {
        pos: live.state.doc.child(0).nodeSize,
        gap: 0,
        height: PAGE,
        breakBefore: false,
        breaks: [],
      },
    ]);
  });

  it("normalizes measurements taken from a visually scaled sheet", () => {
    const live = editor(brokenParagraph());
    live.dom.style.zoom = "0.6";
    draw(live, SHAPES, 0.6);

    expect(measureSheet(live, live.dom).blocks).toEqual([
      { pos: 0, gap: 0, height: 40, breakBefore: false, breaks: [20] },
      {
        pos: live.state.doc.child(0).nodeSize,
        gap: 0,
        height: PAGE,
        breakBefore: false,
        breaks: [],
      },
    ]);
  });

  /**
   * The one thing that cannot be got wrong: the sheet is remeasured on the resize the marks
   * themselves cause, so a measurement that read its own space or push back in would open a
   * wider one every pass and the page would creep away down the sheet.
   */
  it("comes to the same answer once its own marks are on the sheet", () => {
    const live = editor(brokenParagraph());
    draw(live, SHAPES);

    const first = measureSheet(live, live.dom).blocks;
    const applied = layoutOf(first);
    expect(applied.spaces).toHaveLength(1);
    expect(applied.pushes).toHaveLength(1);

    setPagePushes(live, applied.pushes);
    setPageBreakSpaces(live, applied.spaces);
    draw(live, SHAPES);

    const again = measureSheet(live, live.dom).blocks;
    expect(again).toEqual(first);
    expect(layoutOf(again)).toEqual(applied);
  });

  it("leaves a break inside a table to the block it sits in", () => {
    const cell = docxSchema.nodes.tableCell.create(null, [
      docxSchema.nodes.paragraph.create({}, [
        docxSchema.text("in a cell"),
        pageBreak(),
      ]),
    ]);
    const live = editor(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.table.create(null, [
          docxSchema.nodes.tableRow.create(null, [cell]),
        ]),
        docxSchema.nodes.paragraph.create({}, [docxSchema.text("below")]),
      ])
    );
    // A space inside a cell would only grow the cell, so none is put there
    expect(spaceElements(live.dom)).toHaveLength(0);

    draw(live, [
      { gap: 0, height: 40, breaks: [] },
      { gap: 0, height: 20, breaks: [] },
    ]);
    const blocks = measureSheet(live, live.dom).blocks;
    expect(blocks[0]?.breaks).toEqual([]);
    // The break still starts a new page, but only after the whole table
    expect(blocks[1]?.breakBefore).toBe(true);
  });
});
