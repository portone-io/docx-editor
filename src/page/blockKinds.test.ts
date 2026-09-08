// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { Decoration, EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it, vi } from "vitest";
import { docxSchema } from "../schema";
import { editorAttributes } from "../styles/classNames";
import { type BlockKind, blockKindFor } from "./blockKinds";
import { DEFAULT_BLOCK_KINDS } from "./kinds";
import { measureSheet } from "./measureBlocks";
import { pageDecorations, setPageMarks } from "./pageDecorations";
import { A4_PAGE_PIXELS, pageLayout } from "./pageLayout";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  document.body.replaceChildren();
});

function kind(name: string, matches: (node: PMNode) => boolean): BlockKind {
  return {
    name,
    matches,
    measure: () => ({
      candidates: [],
      minFirstPiece: 0,
      breakAfter: false,
      appliedHeight: 0,
    }),
    holdsCut: () => false,
    decorate: () => undefined,
  };
}

function paragraph(...children: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create({}, children);
}

function table(): PMNode {
  return docxSchema.nodes.table.create({ gridCols: [1000] }, [
    docxSchema.nodes.tableRow.create({}, [
      docxSchema.nodes.tableCell.create({}, [paragraph(docxSchema.text("A"))]),
    ]),
  ]);
}

describe("blockKindFor", () => {
  it("the first matching kind wins and the last kind matches every block", () => {
    const anything = kind("anything", () => true);
    const tables = kind(
      "tables",
      (node) => node.type.spec.tableRole === "table"
    );

    expect(blockKindFor([tables, anything], table()).name).toBe("tables");
    expect(blockKindFor([anything, tables], table()).name).toBe("anything");
    expect(
      blockKindFor([tables, anything], paragraph(docxSchema.text("a"))).name
    ).toBe("anything");

    // A registry whose last kind claims nothing leaves a block with no measurer at all, which
    // is a registration mistake rather than a block the engine may skip
    expect(() =>
      blockKindFor([tables], paragraph(docxSchema.text("a")))
    ).toThrow("no block kind matches a paragraph block");
  });
});

/**
 * A shape none of the registered kinds knows: a paragraph parted at an ordinary line break, which
 * the page engine otherwise never cuts at. The kind reads and writes its own attribute, so what
 * it takes to page a new shape is the kind alone.
 */
const LINE_SPACE = "data-test-line-space";

function lineBreaksIn(
  block: PMNode,
  blockPos: number
): { at: number; size: number }[] {
  const found: { at: number; size: number }[] = [];
  block.forEach((child, offset) => {
    if (child.type === docxSchema.nodes.hardBreak && !child.attrs.brAttrs) {
      found.push({ at: blockPos + 1 + offset, size: child.nodeSize });
    }
  });
  return found;
}

const lineBreakKind: BlockKind = {
  name: "line-broken paragraph",

  matches: (node) =>
    node.type === docxSchema.nodes.paragraph &&
    lineBreaksIn(node, 0).length > 0,

  measure({ node, pos, dom, sheetY, top, scale }) {
    const breaks = lineBreaksIn(node, pos);
    const candidates = Array.from(
      dom.querySelectorAll(`[${LINE_SPACE}]`),
      (space, index) => {
        const box = space.getBoundingClientRect();
        return {
          at: breaks[index]?.at ?? 0,
          offset: sheetY(box.top) - top,
          forced: true,
          repeatHeight: 0,
        };
      }
    );
    return {
      candidates,
      minFirstPiece: candidates[0]?.offset ?? 0,
      breakAfter: false,
      appliedHeight:
        Array.from(dom.querySelectorAll(`[${LINE_SPACE}]`)).reduce(
          (total, space) => total + space.getBoundingClientRect().height,
          0
        ) / scale,
    };
  },

  holdsCut: (doc, at) => doc.nodeAt(at)?.type === docxSchema.nodes.hardBreak,

  decorate(pos, node, cuts, into) {
    const heights = new Map(cuts.map((cut) => [cut.at, cut.height]));
    for (const { at, size } of lineBreaksIn(node, pos)) {
      const height = heights.get(at) ?? 0;
      into.push(
        Decoration.inline(at, at + size, {
          nodeName: "span",
          style: `display:block;height:${height}px`,
          [LINE_SPACE]: `${height}`,
        })
      );
    }
  },
};

describe("a kind the editor was built with", () => {
  const PAGE = 500;

  function lineBreak(): PMNode {
    return docxSchema.nodes.hardBreak.create({});
  }

  function mounted(custom: BlockKind = lineBreakKind): EditorView {
    const mount = document.createElement("div");
    document.body.append(mount);
    view = new EditorView(mount, {
      state: EditorState.create({
        doc: docxSchema.nodes.doc.create(null, [
          paragraph(
            docxSchema.text("aaa"),
            lineBreak(),
            docxSchema.text("bbb")
          ),
          paragraph(docxSchema.text("below")),
        ]),
        plugins: [pageDecorations([custom, ...DEFAULT_BLOCK_KINDS])],
      }),
    });
    return view;
  }

  function pixels(value: string | null): number {
    const parsed = Number.parseFloat(value ?? "");
    return Number.isFinite(parsed) ? parsed : 0;
  }

  /**
   * The sheet as the browser would draw it: a block 40 tall whose break stands 20 down, and one
   * a whole page tall below it, with whatever the engine has opened up on them.
   */
  function draw(live: EditorView): void {
    let y = 0;
    live.state.doc.forEach((_node, offset, index) => {
      const dom = live.nodeDOM(offset);
      if (!(dom instanceof HTMLElement)) return;
      const top = y + pixels(dom.getAttribute(editorAttributes.pagePush));
      const space = dom.querySelector(`[${LINE_SPACE}]`);
      const opened = pixels(space?.getAttribute(LINE_SPACE) ?? null);
      if (space) {
        space.getBoundingClientRect = () =>
          new DOMRect(0, top + 20, 400, opened);
      }
      const height = (index === 0 ? 40 : PAGE) + opened;
      dom.getBoundingClientRect = () => new DOMRect(0, top, 400, height);
      y = top + height;
    });
  }

  it("asks the registered kind whether a mapped cut still belongs", () => {
    const holdsCut = vi.fn(lineBreakKind.holdsCut);
    const live = mounted({ ...lineBreakKind, holdsCut });
    setPageMarks(live, { pushes: [], cuts: [{ at: 4, height: 80 }] });

    live.dispatch(live.state.tr.insertText("x", 1));
    expect(holdsCut).toHaveBeenLastCalledWith(live.state.doc, 5);
    expect(
      live.dom.querySelector(`[${LINE_SPACE}]`)?.getAttribute(LINE_SPACE)
    ).toBe("80");

    holdsCut.mockReturnValue(false);
    live.dispatch(live.state.tr.insertText("y", 1));
    expect(holdsCut).toHaveBeenLastCalledWith(live.state.doc, 6);
    expect(
      live.dom.querySelector(`[${LINE_SPACE}]`)?.getAttribute(LINE_SPACE)
    ).toBe("0");
  });

  it("a kind registered for a custom block type measures it and the layout cuts at its candidate", () => {
    const live = mounted();
    const breakAt = 4;
    // The kind drew its own space before anything was measured, which is what the measurement
    // then reads the break's place off
    expect(live.dom.querySelectorAll(`[${LINE_SPACE}]`)).toHaveLength(1);

    draw(live);
    const blocks = measureSheet(live, live.dom).blocks;
    expect(blocks[0]?.candidates).toEqual([
      { at: breakAt, offset: 20, forced: true, repeatHeight: 0 },
    ]);
    // The block below it holds no line break, so it fell through to the paragraph kind
    expect(blocks[1]?.candidates).toEqual([]);

    const layout = pageLayout({
      blocks,
      sections: [
        {
          untilPos: Number.POSITIVE_INFINITY,
          pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: 100 },
        },
      ],
    });
    expect(layout.cuts.map((cut) => cut.at)).toEqual([breakAt]);

    setPageMarks(live, { pushes: layout.pushes, cuts: layout.cuts });
    expect(
      live.dom.querySelector(`[${LINE_SPACE}]`)?.getAttribute(LINE_SPACE)
    ).toBe(`${layout.cuts[0]?.height}`);

    draw(live);
    expect(measureSheet(live, live.dom).blocks).toEqual(blocks);
  });
});
