// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { createEditorState } from "../../editor/createEditor";
import { docxSchema } from "../../schema";
import { editorAttributes } from "../../styles/classNames";
import type { MeasureTarget } from "../blockKinds";
import { paragraphKind } from "./paragraphKind";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  document.body.replaceChildren();
});

function mounted(doc: PMNode): EditorView {
  const mount = document.createElement("div");
  document.body.append(mount);
  view = new EditorView(mount, { state: createEditorState(doc) });
  return view;
}

function pageBreak(): PMNode {
  return docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' });
}

function rect(element: Element, top: number, height: number): void {
  element.getBoundingClientRect = () => new DOMRect(0, top, 400, height);
}

function spaces(live: EditorView): HTMLElement[] {
  return Array.from(
    live.dom.querySelectorAll(`[${editorAttributes.pageBreakSpace}]`)
  );
}

/** The first block of the document, as the sheet hands it to a kind */
function firstBlock(live: EditorView, scale = 1): MeasureTarget {
  const dom = live.nodeDOM(0);
  if (!(dom instanceof HTMLElement)) throw new Error("block DOM not found");
  const sheetY = (viewportY: number) => viewportY / scale;
  return {
    view: live,
    node: live.state.doc.child(0),
    pos: 0,
    dom,
    sheetY,
    top: sheetY(dom.getBoundingClientRect().top),
    scale,
  };
}

/** A paragraph broken twice, so the second break stands below the space the first one opened */
function twiceBroken(): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create({}, [
      docxSchema.text("aaa"),
      pageBreak(),
      docxSchema.text("bbb"),
      pageBreak(),
      docxSchema.text("ccc"),
    ]),
  ]);
}

/** The paragraph 130 tall, its breaks 20 and 90 down, the first with a space of 30 opened at it */
function drawTwiceBroken(live: EditorView, scale = 1): void {
  const dom = live.nodeDOM(0);
  if (!(dom instanceof HTMLElement)) throw new Error("block DOM not found");
  rect(dom, 0, 130 * scale);
  const [first, second] = spaces(live);
  if (!first || !second) throw new Error("space elements not found");
  rect(first, 20 * scale, 30 * scale);
  rect(second, 90 * scale, 0);
}

describe("paragraphKind", () => {
  /**
   * A space is measured for one break and would be a page-sized height on any other, so the
   * measurement has to name the very `br` the element it read belongs to.
   */
  it("pairs each space element with the page br at the same ordinal", () => {
    const live = mounted(twiceBroken());
    // The second break is drawn 30 lower than the place it is measured at, since the space
    // opened at the first one stands between them
    drawTwiceBroken(live);

    expect(paragraphKind.measure(firstBlock(live))).toEqual({
      candidates: [
        { at: 4, offset: 20, forced: true, repeatHeight: 0 },
        { at: 8, offset: 60, forced: true, repeatHeight: 0 },
      ],
      minFirstPiece: 20,
      breakAfter: false,
      appliedHeight: 30,
      keepWithNext: false,
    });
  });

  it("normalizes the spaces of a visually scaled sheet", () => {
    const live = mounted(twiceBroken());
    drawTwiceBroken(live, 0.6);

    expect(paragraphKind.measure(firstBlock(live, 0.6))).toEqual({
      candidates: [
        { at: 4, offset: 20, forced: true, repeatHeight: 0 },
        { at: 8, offset: 60, forced: true, repeatHeight: 0 },
      ],
      minFirstPiece: 20,
      breakAfter: false,
      appliedHeight: 30,
      keepWithNext: false,
    });
  });

  /**
   * The measurement reads where a break stands off the space element itself, so the element has
   * to be on the sheet before there is anything to measure.
   */
  it("puts an empty space on every page br before any measurement", () => {
    const live = mounted(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create({}, [
          docxSchema.text("aaa"),
          pageBreak(),
          docxSchema.text("bbb"),
          pageBreak(),
        ]),
      ])
    );

    expect(
      spaces(live).map((space) => [
        space.tagName,
        space.style.display,
        space.style.height,
        space.querySelectorAll("br").length,
      ])
    ).toEqual([
      ["SPAN", "block", "0px", 1],
      ["SPAN", "block", "0px", 1],
    ]);
  });

  /**
   * A space inside a cell would grow the cell rather than the page, so none is opened there and
   * the fallback rule hands the break to the block it stands in: the next block starts a page.
   */
  it("a br inside a cell is not a paragraph cut", () => {
    const live = mounted(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.table.create({ gridCols: [1000] }, [
          docxSchema.nodes.tableRow.create({}, [
            docxSchema.nodes.tableCell.create({}, [
              docxSchema.nodes.paragraph.create({}, [
                docxSchema.text("in a cell"),
                pageBreak(),
              ]),
            ]),
          ]),
        ]),
      ])
    );
    expect(spaces(live)).toHaveLength(0);

    const measured = paragraphKind.measure(firstBlock(live));
    expect(measured.candidates).toEqual([]);
    expect(measured.breakAfter).toBe(true);
  });

  /**
   * The keep is a display value the formatting resolver derived from the paragraph's own
   * properties or its style, so the kind reads it there and never opens the XML itself.
   */
  it("reads keepNext off the paragraph format", () => {
    const keptWithNext = (pPr: string | null): boolean => {
      const live = mounted(
        docxSchema.nodes.doc.create(null, [
          docxSchema.nodes.paragraph.create({ pPr }, [docxSchema.text("aaa")]),
        ])
      );
      const { keepWithNext } = paragraphKind.measure(firstBlock(live));
      live.destroy();
      view = null;
      return keepWithNext === true;
    };

    expect(keptWithNext("<w:pPr><w:keepNext/></w:pPr>")).toBe(true);
    expect(keptWithNext('<w:pPr><w:keepNext w:val="0"/></w:pPr>')).toBe(false);
    expect(keptWithNext(null)).toBe(false);
  });
});
