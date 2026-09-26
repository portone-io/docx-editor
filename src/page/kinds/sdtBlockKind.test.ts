// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { createEditorState } from "../../editor/createEditor";
import { docxSchema } from "../../schema";
import { editorAttributes } from "../../styles/classNames";
import type { KindMeasure, MeasureTarget, PageCut } from "../blockKinds";
import { DEFAULT_BLOCK_KINDS } from "./index";
import { sdtBlockKind } from "./sdtBlockKind";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  document.body.replaceChildren();
});

const kind = sdtBlockKind;

function paragraph(text: string, attrs: Record<string, unknown> = {}): PMNode {
  return docxSchema.nodes.paragraph.create(attrs, [docxSchema.text(text)]);
}

function pageBreak(): PMNode {
  return docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' });
}

function control(...blocks: PMNode[]): PMNode {
  return docxSchema.nodes.sdtBlock.create(
    { sdtPrefix: "<w:sdt><w:sdtPr/>", key: 1 },
    blocks
  );
}

function table(...rows: PMNode[]): PMNode {
  return docxSchema.nodes.table.create({ gridCols: [1000] }, rows);
}

function row(text: string): PMNode {
  return docxSchema.nodes.tableRow.create({}, [
    docxSchema.nodes.tableCell.create({}, [paragraph(text)]),
  ]);
}

function mounted(block: PMNode): EditorView {
  const mount = document.createElement("div");
  document.body.append(mount);
  view = new EditorView(mount, {
    state: createEditorState(docxSchema.nodes.doc.create(null, [block])),
  });
  return view;
}

function rect(element: Element, top: number, height: number): void {
  element.getBoundingClientRect = () => new DOMRect(0, top, 400, height);
}

/** The block DOM at a position, which the measurement reads every length off */
function drawn(live: EditorView, pos: number): HTMLElement {
  const dom = live.nodeDOM(pos);
  if (!(dom instanceof HTMLElement)) throw new Error(`no DOM at ${pos}`);
  return dom;
}

/** The mounted control as the sheet hands it to its kind */
function target(live: EditorView, scale = 1): MeasureTarget {
  const dom = drawn(live, 0);
  const sheetY = (viewportY: number) => viewportY / scale;
  return {
    view: live,
    node: live.state.doc.child(0),
    pos: 0,
    dom,
    kinds: DEFAULT_BLOCK_KINDS,
    sheetY,
    top: sheetY(dom.getBoundingClientRect().top),
    scale,
  };
}

function measure(live: EditorView, scale = 1): KindMeasure {
  return kind.measure(target(live, scale));
}

/**
 * Lays the control and the blocks it holds out at the heights given, in document order, each one
 * starting where the one before it ended.
 */
function stack(live: EditorView, heights: readonly number[]): number[] {
  const positions: number[] = [];
  let top = 0;
  live.state.doc.child(0).forEach((_child, offset) => {
    positions.push(1 + offset);
  });
  positions.forEach((pos, index) => {
    rect(drawn(live, pos), top, heights[index] ?? 0);
    top += heights[index] ?? 0;
  });
  rect(drawn(live, 0), 0, top);
  return positions;
}

describe("sdtBlockKind", () => {
  it("claims a block control and leaves every other block alone", () => {
    expect(kind.matches(control(paragraph("held")))).toBe(true);
    expect(kind.matches(paragraph("loose"))).toBe(false);
    expect(kind.matches(table(row("cell")))).toBe(false);
  });

  /**
   * The control draws no box of its own, so it stands as tall as the blocks inside it and a page
   * may open between any two of them, exactly as it may between two blocks of the body.
   */
  it("opens a candidate between the blocks it holds", () => {
    const live = mounted(
      control(paragraph("first"), paragraph("second"), paragraph("third"))
    );
    const [, second, third] = stack(live, [40, 50, 60]);

    expect(measure(live)).toEqual({
      candidates: [
        { at: second, offset: 40, forced: false, repeatHeight: 0 },
        { at: third, offset: 90, forced: false, repeatHeight: 0 },
      ],
      opened: new Map(),
      minFirstPiece: 40,
      breakAfter: false,
      breakBefore: false,
      appliedHeight: 0,
      keepWithNext: false,
    });
  });

  /** A page `br` inside a control is the paragraph's own cut, offset into the control */
  it("carries a held paragraph's page break up as its own candidate", () => {
    const live = mounted(
      control(
        paragraph("first"),
        docxSchema.nodes.paragraph.create({}, [
          docxSchema.text("aaa"),
          pageBreak(),
          docxSchema.text("bbb"),
        ])
      )
    );
    const [, second] = stack(live, [40, 60]);
    const space = drawn(live, 0).querySelector(
      `[${editorAttributes.pageBreakSpace}]`
    );
    if (!(space instanceof HTMLElement)) throw new Error("no space element");
    rect(space, 65, 0);

    const measured = measure(live);
    expect(measured.candidates).toEqual([
      { at: second, offset: 40, forced: false, repeatHeight: 0 },
      // The break stands 25 down its own paragraph, which starts 40 down the control
      { at: second + 4, offset: 65, forced: true, repeatHeight: 0 },
    ]);
    expect(measured.minFirstPiece).toBe(40);
  });

  /** A table inside a control keeps its rows, which is what it loses when the control is an atom */
  it("carries a held table's rows up as its own candidates", () => {
    const live = mounted(
      control(paragraph("above"), table(row("one"), row("two"), row("three")))
    );
    const [, tablePos] = stack(live, [40, 150]);
    let rowTop = 40;
    const rowPositions: number[] = [];
    live.state.doc.nodeAt(tablePos)?.forEach((_row, offset) => {
      const pos = tablePos + 1 + offset;
      rowPositions.push(pos);
      rect(drawn(live, pos), rowTop, 50);
      rowTop += 50;
    });

    expect(measure(live).candidates).toEqual([
      { at: tablePos, offset: 40, forced: false, repeatHeight: 0 },
      { at: rowPositions[1], offset: 90, forced: false, repeatHeight: 0 },
      { at: rowPositions[2], offset: 140, forced: false, repeatHeight: 0 },
    ]);
  });

  /** Nesting is what the tree says, so an inner control is paged by this very kind one level down */
  it("reaches the blocks of a control held inside a control", () => {
    const live = mounted(
      control(
        paragraph("outer"),
        control(paragraph("inner"), paragraph("also"))
      )
    );
    const [, innerPos] = stack(live, [40, 100]);
    const inner = live.state.doc.nodeAt(innerPos);
    if (!inner) throw new Error("no inner control");
    let top = 40;
    const innerPositions: number[] = [];
    inner.forEach((_child, offset) => {
      const pos = innerPos + 1 + offset;
      innerPositions.push(pos);
      rect(drawn(live, pos), top, 50);
      top += 50;
    });

    expect(measure(live).candidates).toEqual([
      { at: innerPos, offset: 40, forced: false, repeatHeight: 0 },
      { at: innerPositions[1], offset: 90, forced: false, repeatHeight: 0 },
    ]);
  });

  /**
   * Every length is read off the block it belongs to, so the control never runs a query over its
   * own subtree and a paragraph is not scanned again for each control standing around it.
   */
  it("reads no query over the whole control", () => {
    const live = mounted(
      control(
        paragraph("first"),
        docxSchema.nodes.paragraph.create({}, [
          docxSchema.text("aaa"),
          pageBreak(),
        ])
      )
    );
    stack(live, [40, 60]);
    const dom = drawn(live, 0);
    const queries: string[] = [];
    const original = dom.querySelectorAll.bind(dom);
    dom.querySelectorAll = (selector: string) => {
      queries.push(selector);
      return original(selector);
    };

    measure(live);
    expect(queries).toEqual([]);
  });

  /** A paragraph that asks for a page of its own is a cut the control has to make */
  it("forces a candidate at a held paragraph that opens a page", () => {
    const live = mounted(
      control(
        paragraph("first"),
        paragraph("second", { pPr: "<w:pPr><w:pageBreakBefore/></w:pPr>" })
      )
    );
    const [, second] = stack(live, [40, 50]);

    expect(measure(live).candidates).toEqual([
      { at: second, offset: 40, forced: true, repeatHeight: 0 },
    ]);
  });

  /**
   * A control draws nothing of its own, so a page the first block inside it asks for is a page the
   * sheet can only hear about from the kind (`page/measureBlocks`).
   */
  it("reports a page the first block it holds asks for", () => {
    const live = mounted(
      control(
        paragraph("first", { pPr: "<w:pPr><w:pageBreakBefore/></w:pPr>" })
      )
    );
    stack(live, [40]);

    const measured = measure(live);
    expect(measured.breakBefore).toBe(true);
    expect(measured.candidates).toEqual([]);
  });

  /**
   * A space cannot be opened inside a cell, so the break a held table carries cuts at the block
   * after it; only a break under the last block is left for the sheet to answer after the control.
   */
  it("answers a break it could not open a space at with the next block", () => {
    const withBreakInACell = table(
      docxSchema.nodes.tableRow.create({}, [
        docxSchema.nodes.tableCell.create({}, [
          docxSchema.nodes.paragraph.create({}, [
            docxSchema.text("in a cell"),
            pageBreak(),
          ]),
        ]),
      ])
    );
    const live = mounted(control(withBreakInACell, paragraph("after")));
    const [, after] = stack(live, [60, 40]);

    const measured = measure(live);
    expect(measured.candidates).toEqual([
      { at: after, offset: 60, forced: true, repeatHeight: 0 },
    ]);
    expect(measured.breakAfter).toBe(false);

    const last = mounted(control(paragraph("before"), withBreakInACell));
    stack(last, [40, 60]);
    expect(measure(last).breakAfter).toBe(true);
  });

  /** The keep the control asks for is the one its last block asks for */
  it("keeps with the next block where the last block it holds does", () => {
    const live = mounted(
      control(
        paragraph("first", { pPr: "<w:pPr><w:keepNext/></w:pPr>" }),
        paragraph("last")
      )
    );
    stack(live, [40, 40]);
    expect(measure(live).keepWithNext).toBe(false);

    const kept = mounted(
      control(
        paragraph("first"),
        paragraph("last", { pPr: "<w:pPr><w:keepNext/></w:pPr>" })
      )
    );
    stack(kept, [40, 40]);
    expect(measure(kept).keepWithNext).toBe(true);
  });

  /**
   * A control draws none of the properties of the blocks it holds, its own div included, so a page
   * the first block of a control held inside one asks for reaches the sheet only as the answer the
   * inner kind gives (`page/blockKinds.opensPage`).
   */
  it("reports a page the first block of a control held inside it asks for", () => {
    const live = mounted(
      control(
        control(
          paragraph("first", { pPr: "<w:pPr><w:pageBreakBefore/></w:pPr>" })
        )
      )
    );
    const innerPos = 1;
    rect(drawn(live, innerPos), 0, 40);
    rect(drawn(live, innerPos + 1), 0, 40);
    rect(drawn(live, 0), 0, 40);

    expect(measure(live).breakBefore).toBe(true);
  });

  it("forces the boundary at a held control that opens a page", () => {
    const live = mounted(
      control(
        paragraph("loose"),
        control(
          paragraph("held", { pPr: "<w:pPr><w:pageBreakBefore/></w:pPr>" })
        )
      )
    );
    const [, innerPos] = stack(live, [40, 50]);
    rect(drawn(live, innerPos + 1), 40, 50);

    expect(measure(live).candidates).toEqual([
      { at: innerPos, offset: 40, forced: true, repeatHeight: 0 },
    ]);
  });

  /**
   * What has to fit on the page the control starts on is what its first block asks for, so a held
   * table still asks for its repeated headers followed by one body row rather than for the
   * boundary right under the header.
   */
  it("asks for the first piece its first block asks for", () => {
    const header = docxSchema.nodes.tableRow.create(
      { format: { repeatHeader: true } },
      [docxSchema.nodes.tableCell.create({}, [paragraph("Heading")])]
    );
    const live = mounted(
      control(table(header, row("one"), row("two")), paragraph("after"))
    );
    const [tablePos] = stack(live, [150, 40]);
    let rowTop = 0;
    live.state.doc.nodeAt(tablePos)?.forEach((_row, offset) => {
      rect(drawn(live, tablePos + 1 + offset), rowTop, 50);
      rowTop += 50;
    });

    // The headers plus one body row, not the boundary under the header alone
    expect(measure(live).minFirstPiece).toBe(100);
  });

  /**
   * A keep (§17.3.1.14) closes the boundary under the block asking for it, which the layout
   * opens only when no page can hold the two blocks kept together (`page/pageLayout`).
   */
  it("offers the boundary under a block kept with the one after it as kept", () => {
    const live = mounted(
      control(
        paragraph("first", { pPr: "<w:pPr><w:keepNext/></w:pPr>" }),
        paragraph("second")
      )
    );
    const [, second] = stack(live, [40, 50]);

    expect(measure(live).candidates).toEqual([
      { at: second, offset: 40, forced: false, kept: true, repeatHeight: 0 },
    ]);
  });

  /** A page the document asks for beats the keep, exactly as it does between two body blocks */
  it("offers a forced boundary under a kept block all the same", () => {
    const live = mounted(
      control(
        paragraph("first", { pPr: "<w:pPr><w:keepNext/></w:pPr>" }),
        paragraph("second", { pPr: "<w:pPr><w:pageBreakBefore/></w:pPr>" })
      )
    );
    const [, second] = stack(live, [40, 50]);

    expect(measure(live).candidates).toEqual([
      { at: second, offset: 40, forced: true, repeatHeight: 0 },
    ]);
  });

  describe("the cuts it was given", () => {
    it("opens a space between two of the blocks it holds", () => {
      const live = mounted(control(paragraph("first"), paragraph("second")));
      const [, second] = stack(live, [40, 50]);
      const cuts: PageCut[] = [{ at: second, height: 120 }];

      const decorations: Parameters<typeof kind.decorate>[3] = [];
      kind.decorate(
        0,
        live.state.doc.child(0),
        cuts,
        decorations,
        DEFAULT_BLOCK_KINDS
      );
      const space = decorations
        .map((decoration) => decoration.spec)
        .filter((spec) => typeof spec.key === "string");
      expect(space).toHaveLength(1);
      expect(space[0]?.key).toBe(`container-page-space-${second}-120`);
    });

    it("hands a cut inside one of them to that block's own kind", () => {
      const live = mounted(
        control(
          paragraph("first"),
          docxSchema.nodes.paragraph.create({}, [
            docxSchema.text("aaa"),
            pageBreak(),
          ])
        )
      );
      const [, second] = stack(live, [40, 50]);
      const at = second + 4;
      const decorations: Parameters<typeof kind.decorate>[3] = [];
      kind.decorate(
        0,
        live.state.doc.child(0),
        [{ at, height: 90 }],
        decorations,
        DEFAULT_BLOCK_KINDS
      );

      // The paragraph draws its own break space, at the height the cut asked for
      expect(
        decorations.map((decoration) => [decoration.from, decoration.to])
      ).toContainEqual([at, at + 1]);
    });

    it("holds a boundary between its blocks and a place inside one of them", () => {
      const live = mounted(
        control(
          paragraph("first"),
          docxSchema.nodes.paragraph.create({}, [
            docxSchema.text("aaa"),
            pageBreak(),
          ])
        )
      );
      const [first, second] = stack(live, [40, 50]);
      const { doc } = live.state;

      expect(kind.holdsCut(doc, second, DEFAULT_BLOCK_KINDS)).toBe(true);
      expect(kind.holdsCut(doc, second + 4, DEFAULT_BLOCK_KINDS)).toBe(true);
      // The first block is where the control starts, so a cut there parts nothing
      expect(kind.holdsCut(doc, first, DEFAULT_BLOCK_KINDS)).toBe(false);
      expect(kind.holdsCut(doc, second + 1, DEFAULT_BLOCK_KINDS)).toBe(false);
    });

    it("holds a boundary inside a control held inside it", () => {
      const live = mounted(
        control(
          paragraph("outer"),
          control(paragraph("inner"), paragraph("also"))
        )
      );
      const [, innerPos] = stack(live, [40, 100]);
      const { doc } = live.state;
      const inner = doc.nodeAt(innerPos);
      if (!inner) throw new Error("no inner control");

      expect(
        kind.holdsCut(
          doc,
          innerPos + 1 + inner.child(0).nodeSize,
          DEFAULT_BLOCK_KINDS
        )
      ).toBe(true);
      expect(kind.holdsCut(doc, innerPos + 1, DEFAULT_BLOCK_KINDS)).toBe(false);
    });
  });
});

/**
 * A shut control paints a background over everything it holds, and the space a cut opens is one
 * of its own children, so the fill would run from the cut down to the foot of the page. Only the
 * CSS breaks it.
 */
describe("the space rule in editor.css", () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "../../styles/editor.css"),
    "utf8"
  );

  /** What the rule a selector opens declares, the selector read as the text it is */
  function declarations(selector: string): string {
    const literal = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const found = css.match(new RegExp(`\\n${literal}\\s*\\{([^}]*)\\}`))?.[1];
    if (found === undefined) throw new Error(`rule not found: ${selector}`);
    return found;
  }

  function background(selector: string): string {
    const found = declarations(selector).match(/background:\s*([^;]+);/)?.[1];
    if (found === undefined) throw new Error(`no background: ${selector}`);
    return found.trim();
  }

  it("paints the paper back over the fill a shut control lays down", () => {
    expect(declarations(".docx-editor-sdt-locked")).toMatch(
      /background-image:\s*linear-gradient\(\s*var\(--docx-editor-locked-background/
    );
    expect(
      background(
        `.docx-editor-sdt-block > [${editorAttributes.containerPageSpace}]`
      )
    ).toBe(background(".docx-editor-sheet"));
  });
});
