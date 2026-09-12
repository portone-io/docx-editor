// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { editorAttributes } from "../styles/classNames";
import type { DemandSource } from "./demands";
import { DEFAULT_BLOCK_KINDS } from "./kinds";
import { measureSheet } from "./measureBlocks";
import { pageDecorations, setPageMarks } from "./pageDecorations";
import { A4_PAGE_PIXELS, type MeasuredBlock, pageLayout } from "./pageLayout";

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

/** Where the page break in the first block stands, as the measurement names it */
function breakPos(live: EditorView): number {
  let found = -1;
  live.state.doc.child(0).forEach((child, offset) => {
    if (child.type === docxSchema.nodes.hardBreak) found = 1 + offset;
  });
  return found;
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
  scale = 1,
  /**
   * What a space comes out taller than it was given, the way a table's spacer row takes half of a
   * collapsed border on each side
   */
  overshoot = 0
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
      const asked = number(
        element.getAttribute(editorAttributes.pageBreakSpace)
      );
      const height = asked > 0 ? asked + overshoot : asked;
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
  return pageLayout({
    blocks,
    sections: [
      {
        untilPos: Number.POSITIVE_INFINITY,
        pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: STEP },
        type: null,
      },
    ],
  });
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

/** `SHAPES` as the blocks a measurement of them comes to */
function measuredShapes(live: EditorView): MeasuredBlock[] {
  return [
    {
      pos: 0,
      gap: 0,
      height: 40,
      breakBefore: false,
      breakAfter: false,
      candidates: [
        { at: breakPos(live), offset: 20, forced: true, repeatHeight: 0 },
      ],
      minFirstPiece: 20,
      keepWithNext: false,
      demands: [],
    },
    {
      pos: live.state.doc.child(0).nodeSize,
      gap: 0,
      height: PAGE,
      breakBefore: false,
      breakAfter: false,
      candidates: [],
      minFirstPiece: PAGE,
      keepWithNext: false,
      demands: [],
    },
  ];
}

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
    expect(measureSheet(live, live.dom).blocks).toEqual(measuredShapes(live));
  });

  it("normalizes measurements taken from a visually scaled sheet", () => {
    const live = editor(brokenParagraph());
    live.dom.style.zoom = "0.6";
    draw(live, SHAPES, 0.6);

    expect(measureSheet(live, live.dom).blocks).toEqual(measuredShapes(live));
  });

  /**
   * The one thing that cannot be got wrong: the sheet is remeasured on the resize the marks
   * themselves cause, so a measurement that read its own space or push back in would open a
   * wider one every pass and the page would creep away down the sheet.
   */
  it("comes to the same answer once its own marks are on the sheet", () => {
    const live = editor(brokenParagraph());
    draw(live, SHAPES);

    const first = measureSheet(live, live.dom);
    const applied = layoutOf(first.blocks);
    expect(applied.cuts).toHaveLength(1);
    expect(applied.pushes).toHaveLength(1);

    setPageMarks(live, { pushes: applied.pushes, cuts: applied.cuts });
    draw(live, SHAPES);

    const again = measureSheet(live, live.dom);
    expect(again.blocks).toEqual(first.blocks);
    expect(layoutOf(again.blocks)).toEqual(applied);
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
    expect(blocks[0]?.candidates).toEqual([]);
    // The break still starts a new page, but only after the whole table, which the block it
    // sits in reports and the layout answers
    expect(blocks[0]?.breakAfter).toBe(true);
    expect(blocks[1]?.breakBefore).toBe(false);
  });

  it("carries a block's keep with next through to the layout", () => {
    const live = editor(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create(
          { pPr: "<w:pPr><w:keepNext/></w:pPr>" },
          [docxSchema.text("kept")]
        ),
        docxSchema.nodes.paragraph.create({}, [docxSchema.text("next")]),
      ])
    );
    draw(live, [
      { gap: 0, height: 40, breaks: [] },
      { gap: 0, height: 40, breaks: [] },
    ]);

    expect(
      measureSheet(live, live.dom).blocks.map((block) => block.keepWithNext)
    ).toEqual([true, false]);
  });
});

describe("the demands of a block", () => {
  /** An editor holding the default kinds and these demand sources alone */
  function asking(doc: PMNode, sources: readonly DemandSource[]): EditorView {
    const mount = document.createElement("div");
    layer = mount;
    document.body.appendChild(mount);
    view = new EditorView(mount, {
      state: EditorState.create({
        doc,
        plugins: [pageDecorations(DEFAULT_BLOCK_KINDS, sources)],
      }),
    });
    return view;
  }

  it("asks every registered demand source about every block whatever its kind", () => {
    const asked: string[] = [];
    const source = (name: string): DemandSource => ({
      name,
      demandsIn: ({ node, pos }) => {
        asked.push(`${name} ${node.type.name} ${pos}`);
        return [{ offset: 5, id: `${name} ${pos}`, band: "test" }];
      },
    });
    const table = docxSchema.nodes.table.create({ gridCols: [1000] }, [
      docxSchema.nodes.tableRow.create({}, [
        docxSchema.nodes.tableCell.create({}, [
          docxSchema.nodes.paragraph.create({}, [docxSchema.text("in a cell")]),
        ]),
      ]),
    ]);
    const live = asking(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create({}, [docxSchema.text("above")]),
        table,
        docxSchema.nodes.paragraph.create({}, [docxSchema.text("below")]),
      ]),
      [source("first"), source("second")]
    );
    draw(live, [
      { gap: 0, height: 20, breaks: [] },
      { gap: 0, height: 40, breaks: [] },
      { gap: 0, height: 20, breaks: [] },
    ]);

    const tablePos = live.state.doc.child(0).nodeSize;
    const belowPos = tablePos + table.nodeSize;
    const blocks = measureSheet(live, live.dom).blocks;
    expect(asked).toEqual([
      "first paragraph 0",
      "second paragraph 0",
      `first table ${tablePos}`,
      `second table ${tablePos}`,
      `first paragraph ${belowPos}`,
      `second paragraph ${belowPos}`,
    ]);
    expect(
      blocks.map((block) => block.demands?.map((demand) => demand.id))
    ).toEqual([
      ["first 0", "second 0"],
      [`first ${tablePos}`, `second ${tablePos}`],
      [`first ${belowPos}`, `second ${belowPos}`],
    ]);
  });

  /**
   * The height a cut asked for and the height the browser drew it at are not the same number: a
   * table's spacer row takes half of a collapsed border on each side. Everything a kind reports is
   * read off the sheet as drawn, so a place below the cut has to be taken back by what stands
   * there rather than by what was asked for, or it drifts below the piece it belongs to and its
   * room is kept on the wrong page.
   */
  it("reads a place below a cut drawn taller than the cut asked for at its own offset", () => {
    const places: DemandSource = {
      name: "below the break",
      demandsIn: ({ pos, dom, sheetY, top }) =>
        pos === 0
          ? [
              {
                offset: sheetY(dom.getBoundingClientRect().bottom) - top - 10,
                id: "below",
                band: "test",
              },
            ]
          : [],
    };
    const live = asking(brokenParagraph(), [places]);
    draw(live, SHAPES);

    const natural = measureSheet(live, live.dom).blocks[0]?.demands;
    const applied = layoutOf(measureSheet(live, live.dom).blocks);
    expect(applied.cuts).toHaveLength(1);

    setPageMarks(live, { pushes: applied.pushes, cuts: applied.cuts });
    // The browser draws the space three pixels taller than the cut asked for
    draw(live, SHAPES, 1, 3);

    expect(measureSheet(live, live.dom).blocks[0]?.demands).toEqual(natural);
  });

  /**
   * A source reads its places off the sheet as drawn, so a place below a break reads the space the
   * layout opened there. Read in, that space would move the place onto the next page's count, and
   * the page would reserve room for it again on the pass its own marks cause.
   */
  it("reads a place below a cut where it stands in the block with no space opened", () => {
    const places: DemandSource = {
      name: "above and below the break",
      demandsIn: ({ pos, dom, sheetY, top }) =>
        pos === 0
          ? [
              { offset: 10, id: "above", band: "test" },
              {
                offset: sheetY(dom.getBoundingClientRect().bottom) - top - 10,
                id: "below",
                band: "test",
              },
            ]
          : [],
    };
    const live = asking(brokenParagraph(), [places]);
    draw(live, SHAPES);

    const first = measureSheet(live, live.dom);
    expect(first.blocks[0]?.demands).toEqual([
      { offset: 10, id: "above", band: "test" },
      { offset: 30, id: "below", band: "test" },
    ]);

    const applied = layoutOf(first.blocks);
    expect(applied.cuts).toHaveLength(1);
    setPageMarks(live, { pushes: applied.pushes, cuts: applied.cuts });
    draw(live, SHAPES);

    expect(measureSheet(live, live.dom).blocks).toEqual(first.blocks);
  });
});
