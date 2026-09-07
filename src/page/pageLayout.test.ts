// @vitest-environment node
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { LETTER_GEOMETRY } from "../__testing__/docx";
import { editorClassNames, editorCssVariables } from "../styles/classNames";
import type { BreakCandidate } from "./blockKinds";
import {
  A4_PAGE_PIXELS,
  type MeasuredBlock,
  PAGE_SPLIT_PX,
  pageGeometryStyle,
  pageLayout,
  pagePixels,
} from "./pageLayout";

const PAGE = 1000;
const STEP = 200;

/** Blocks laid out one after another with no gap between them */
function blocks(
  ...heights: readonly (number | Partial<MeasuredBlock>)[]
): MeasuredBlock[] {
  return heights.map((entry, index) => {
    const shape = typeof entry === "number" ? { height: entry } : entry;
    const height = shape.height ?? 0;
    return {
      pos: index * 10,
      gap: 0,
      breakBefore: false,
      breakAfter: false,
      candidates: [],
      minFirstPiece: height,
      ...shape,
      height,
    };
  });
}

/** A block parted where the text says it is, the way a paragraph's page breaks are measured */
function broken(
  height: number,
  ...offsets: readonly number[]
): Partial<MeasuredBlock> {
  return {
    height,
    candidates: offsets.map((offset, index) => ({
      at: 100 + index,
      offset,
      forced: true,
      repeatHeight: 0,
    })),
    minFirstPiece: offsets[0] ?? height,
  };
}

/** Row boundaries as a table offers them: taken only where the piece after one would overflow */
function rowBoundaries(
  repeatHeight: number,
  ...offsets: readonly number[]
): BreakCandidate[] {
  return offsets.map((offset, index) => ({
    at: 101 + index,
    offset,
    forced: false,
    repeatHeight,
  }));
}

function layout(list: MeasuredBlock[]) {
  return pageLayout({
    blocks: list,
    pageBodyHeight: PAGE,
    pageStep: STEP,
  });
}

describe("pageLayout", () => {
  it("nothing to push and nothing to split when it all fits on one page", () => {
    const result = layout(blocks(300, 300, 300));
    expect(result.pushes).toEqual([]);
    expect(result.splits).toEqual([]);
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, exactPage: true, crossed: false },
    ]);
    expect(result.bodyHeight).toBe(PAGE);
  });

  it("moves a block that straddles a boundary whole to the top of the next page", () => {
    const result = layout(blocks(900, 300));
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 900, push: PAGE + STEP - 900 },
    ]);
    expect(result.splits).toMatchObject([{ y: PAGE, page: 2, crossed: false }]);
    expect(result.bodyHeight).toBe(PAGE + STEP + PAGE);
  });

  it("the blocks below move down by the same amount and the next boundary is counted again", () => {
    // With 700 following 900 + (the pushed 300), the second page's body does not exceed
    // 1000
    const result = layout(blocks(900, 300, 700));
    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
    expect(result.pages).toHaveLength(2);
  });

  it("keeps the gap the block already had by adding it on top of the pushed value", () => {
    const result = layout(blocks(900, { height: 300, gap: 40 }));
    expect(result.pushes).toEqual([
      {
        pos: 10,
        marginTop: 40 + (PAGE + STEP - 940),
        push: PAGE + STEP - 940,
      },
    ]);
  });

  it("a block the document wrote a break onto goes to the next page even without straddling one", () => {
    const result = layout(blocks(300, { height: 200, breakBefore: true }));
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 300, push: PAGE + STEP - 300 },
    ]);
    expect(result.splits).toMatchObject([{ page: 2, forced: true }]);
  });

  it("a break written on the first block does not make an empty page", () => {
    const result = layout(blocks({ height: 200, breakBefore: true }));
    expect(result.pushes).toEqual([]);
    expect(result.splits).toEqual([]);
  });

  it("a page break inside a block fills out the page, so the rest of the block starts the next one", () => {
    const result = layout(blocks(broken(300, 100)));
    // The 900 left on the page, plus the step over to the next page's body
    expect(result.cuts).toEqual([{ at: 100, height: PAGE - 100 + STEP }]);
    expect(result.pushes).toEqual([]);
    expect(result.splits).toMatchObject([
      { y: PAGE, page: 2, forced: true, crossed: false },
    ]);
    expect(result.bodyHeight).toBe(PAGE + STEP + PAGE);
  });

  it("only the part of a block before its first break has to fit on the page it starts on", () => {
    // The 50 up to the break fits in the 100 left over, where the whole 300 would not
    const result = layout(blocks(900, broken(300, 50)));
    expect(result.pushes).toEqual([]);
    expect(result.cuts).toEqual([{ at: 100, height: PAGE - 950 + STEP }]);
  });

  it("a break at the end of a block puts the block after it at the top of the next page", () => {
    // The second block is a page tall, so it only fits at all where the break left it
    const result = layout(blocks(broken(300, 300), PAGE));
    expect(result.cuts).toEqual([{ at: 100, height: PAGE - 300 + STEP }]);
    expect(result.pushes).toEqual([]);
    expect(result.splits).toHaveLength(1);
    expect(result.bodyHeight).toBe(PAGE + STEP + PAGE);
  });

  it("each break in a block fills out the page it lands on", () => {
    const result = layout(blocks(broken(400, 100, 250)));
    // The second break stands 150 into the page the first one opened, not 250 into the block
    expect(result.cuts).toEqual([
      { at: 100, height: PAGE - 100 + STEP },
      { at: 101, height: PAGE - 150 + STEP },
    ]);
    expect(result.splits.map((split) => split.forced)).toEqual([true, true]);
    expect(result.pages).toHaveLength(3);
  });

  it("a break inside a block taller than a page is answered from the page it stands on", () => {
    const result = layout(blocks(broken(2500, 1500)));
    expect(result.splits).toMatchObject([
      { y: PAGE, page: 2, crossed: true },
      { y: 2 * PAGE, page: 3, forced: true, crossed: false },
    ]);
    expect(result.cuts).toEqual([{ at: 100, height: 2 * PAGE - 1500 + STEP }]);
  });

  it("a page split by a break inside a block has an exact number", () => {
    const result = layout(blocks(broken(300, 100)));
    expect(result.pages.map((start) => start.exactPage)).toEqual([true, true]);
  });

  it("a page split by written breaks alone has an exact number, and becomes an estimate once an estimate is involved", () => {
    const exact = layout(
      blocks(
        200,
        { height: 200, breakBefore: true },
        {
          height: 200,
          breakBefore: true,
        }
      )
    );
    expect(exact.splits.map((split) => split.exactPage)).toEqual([true, true]);

    const mixed = layout(blocks(900, 300, { height: 200, breakBefore: true }));
    expect(mixed.splits.map((split) => split.exactPage)).toEqual([
      false,
      false,
    ]);
  });

  it("reports where each page starts", () => {
    const result = layout(blocks(900, 300));
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, exactPage: true, crossed: false },
      { page: 2, bodyStart: PAGE + STEP, exactPage: false, crossed: false },
    ]);
  });

  it("a page reached by crossing joins onto the previous page with no top margin", () => {
    const result = layout(blocks(1500));
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, exactPage: true, crossed: false },
      { page: 2, bodyStart: PAGE, exactPage: false, crossed: true },
    ]);
  });

  it("does not push a block taller than a page and only reports where it crossed", () => {
    const result = layout(blocks(2500));
    expect(result.pushes).toEqual([]);
    expect(result.splits).toMatchObject([
      { y: PAGE, page: 2, crossed: true },
      { y: 2 * PAGE, page: 3, crossed: true },
    ]);
    // No gap opens up between pages while the block is crossing them
    expect(result.bodyHeight).toBe(3 * PAGE);
  });

  it("the block after a crossing block joins onto the space left on that page", () => {
    const result = layout(blocks(1500, 200));
    expect(result.pushes).toEqual([]);
    expect(result.splits).toHaveLength(1);
    expect(result.pages).toHaveLength(2);
  });

  it("the next block moves down to the page after when a crossing block fills the page", () => {
    const result = layout(blocks(1990, 200));
    expect(result.splits.map((split) => split.crossed)).toEqual([true, false]);
    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
  });

  it("continues a long table at the last row boundary that fits", () => {
    const result = layout(
      blocks({
        height: 1500,
        candidates: rowBoundaries(0, 300, 600, 900, 1200),
        minFirstPiece: 300,
      })
    );

    expect(result.cuts).toEqual([{ at: 103, height: 300 }]);
    expect(result.splits).toMatchObject([{ y: PAGE, crossed: false }]);
  });

  it("moves a long table first when its first row does not fit the current page", () => {
    const table = blocks({
      height: 1200,
      candidates: rowBoundaries(0, 300, 600, 900),
      minFirstPiece: 300,
    })[0];
    if (!table) throw new Error("table block not built");
    const result = layout([...blocks(800), { ...table, pos: 10 }]);

    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
    expect(result.cuts).toEqual([{ at: 103, height: 300 }]);
    expect(result.splits).toHaveLength(2);
  });

  it("repeats table headers after the page space", () => {
    const result = layout(
      blocks({
        height: 1300,
        candidates: rowBoundaries(100, 100, 400, 700, 1000),
        minFirstPiece: 400,
      })
    );

    expect(result.cuts).toEqual([{ at: 104, height: STEP }]);
    expect(result.bodyHeight).toBe(PAGE + STEP + PAGE);
  });

  it("crosses only when a row group between safe boundaries is taller than a page", () => {
    const result = layout(
      blocks({
        height: 1600,
        candidates: rowBoundaries(0, 1300),
        minFirstPiece: 1300,
      })
    );

    expect(result.cuts).toEqual([]);
    expect(result.splits.map((split) => split.crossed)).toEqual([true]);
  });

  it("leaves a candidate uncut when the piece after it is taller than a page", () => {
    const result = layout(
      blocks({
        height: 1600,
        candidates: rowBoundaries(0, 300),
        minFirstPiece: 300,
      })
    );

    // Cutting here would open a space the piece cannot be made to fit behind, and the text would
    // cross the next boundary anyway: one more page for nothing
    expect(result.cuts).toEqual([]);
    expect(result.splits.map((split) => split.crossed)).toEqual([true]);
    expect(result.pages).toHaveLength(2);
    expect(result.bodyHeight).toBe(2 * PAGE);
  });

  it("a forced candidate and an optional one in the same block cut in document order", () => {
    const result = layout(
      blocks({
        height: 1400,
        candidates: [
          { at: 7, offset: 300, forced: true, repeatHeight: 0 },
          { at: 105, offset: 900, forced: false, repeatHeight: 0 },
        ],
        minFirstPiece: 300,
      })
    );

    // The optional cut is answered from where the forced one left the piece before it, not from
    // the block's own top: one running total carries both
    expect(result.cuts).toEqual([
      { at: 7, height: PAGE - 300 + STEP },
      { at: 105, height: 600 },
    ]);
    expect(result.splits.map((split) => split.forced)).toEqual([true, false]);
  });

  it("an optional candidate is answered before the forced one that follows it", () => {
    const result = layout(
      blocks({
        height: 1400,
        candidates: [
          { at: 100, offset: 300, forced: false, repeatHeight: 0 },
          { at: 200, offset: 1200, forced: true, repeatHeight: 0 },
        ],
        minFirstPiece: 300,
      })
    );

    // Reading the block as breaks first and boundaries second would carry the forced cut's space
    // back to a piece that stands before it, leaving the optional one uncut and crossed instead
    expect(result.cuts).toEqual([
      { at: 100, height: 900 },
      { at: 200, height: 300 },
    ]);
    expect(result.splits.map((split) => split.forced)).toEqual([false, true]);
    expect(result.splits.map((split) => split.crossed)).toEqual([false, false]);
  });

  it("an optional candidate with a repeat height carries it onto the next page before the piece", () => {
    const result = layout(
      blocks({
        height: 2000,
        candidates: rowBoundaries(200, 900, 1400),
        minFirstPiece: 900,
      })
    );

    // The 200 of repeated header stands above the continued piece, so the piece starts 200
    // lower and the cut after it opens that much less space
    expect(result.cuts).toEqual([
      { at: 101, height: 300 },
      { at: 102, height: 500 },
    ]);
  });

  it("a block with breakAfter starts the next block on a new page", () => {
    const result = layout(blocks({ height: 200, breakAfter: true }, 200));

    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 200, push: PAGE + STEP - 200 },
    ]);
    expect(result.splits).toMatchObject([{ page: 2, forced: true }]);
  });

  it("does nothing for a page height that cannot be measured", () => {
    expect(
      pageLayout({ blocks: blocks(500), pageBodyHeight: 0, pageStep: STEP })
    ).toMatchObject({ pushes: [], splits: [] });
  });

  it("one page's body height is 25.7cm, the A4 fallback minus its margins", () => {
    // A4 is 29.7cm of paper less 2cm of margin at either end, rounded to whole twips
    expect(A4_PAGE_PIXELS.bodyHeight).toBeCloseTo((25.7 * 96) / 2.54, 1);
    // From the end of the previous page's body to the top of the next page's body is two
    // margins plus the grey gap
    expect(A4_PAGE_PIXELS.pageStep).toBeCloseTo((4 * 96) / 2.54 + 24, 1);
  });

  it("draws a Letter document on Letter, not on the A4 fallback", () => {
    const letter = pagePixels(LETTER_GEOMETRY);
    // 8.5in x 11in at 96px to the inch, with an inch of margin all round
    expect(letter.pageWidth).toBeCloseTo(816, 6);
    expect(letter.pageHeight).toBeCloseTo(1056, 6);
    expect(letter.bodyWidth).toBeCloseTo(816 - 96 * 2, 6);
    expect(letter.bodyHeight).toBeCloseTo(1056 - 96 * 2, 6);
    expect(letter.pageStep).toBeCloseTo(96 * 2 + PAGE_SPLIT_PX, 6);

    // Letter is the wider paper but the shorter one, and its inch of margin leaves a
    // narrower body than A4's 2.2cm does. Neither number is the fallback's
    expect(letter.pageWidth).toBeGreaterThan(A4_PAGE_PIXELS.pageWidth);
    expect(letter.pageHeight).toBeLessThan(A4_PAGE_PIXELS.pageHeight);
    expect(letter.bodyWidth).not.toBeCloseTo(A4_PAGE_PIXELS.bodyWidth, 1);
    expect(letter.bodyHeight).toBeLessThan(A4_PAGE_PIXELS.bodyHeight);
  });

  it("hands the paper to the sheet as the CSS variables it is drawn from", () => {
    const style = pageGeometryStyle(A4_PAGE_PIXELS);
    expect(style).toContain(
      `${editorCssVariables.pageWidth}:${Math.round(A4_PAGE_PIXELS.pageWidth * 100) / 100}px`
    );
    expect(style).toContain(editorCssVariables.pageMarginLeft);
    expect(style).toContain(editorCssVariables.pageHeight);
  });
});

/**
 * The paper an open document is drawn on comes from its own `w:pgSz`, handed to the sheet as
 * CSS variables, so the two cannot drift while a document is open.
 * What `editor.css` writes down is the paper used before any document is opened, and a
 * document that names none is drawn on that same paper - so the CSS has to say what
 * `A4_PORTRAIT` says, or the sheet would jump the moment a plain document loaded.
 */
describe("the paper size in editor.css", () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "../styles/editor.css"),
    "utf8"
  );

  /** One of the values recorded in cm in the CSS */
  function cm(pattern: RegExp): number {
    const found = css.match(pattern)?.[1];
    if (found === undefined) throw new Error(`value not found: ${pattern}`);
    return Number.parseFloat(found);
  }

  /**
   * One of the paper variables declared on the sheet rule itself, so a value written
   * anywhere earlier in the file cannot stand in for it
   */
  function sheetVariable(name: string): RegExp {
    return new RegExp(
      `\\.${editorClassNames.sheet}\\s*\\{[^}]*?${name}:\\s*([\\d.]+)cm`
    );
  }

  const PX_PER_CM = 96 / 2.54;
  const pageWidth = cm(sheetVariable(editorCssVariables.pageWidth));
  const pageHeight = cm(sheetVariable(editorCssVariables.pageHeight));
  const margin = cm(sheetVariable(editorCssVariables.pageMarginTop));
  const sideMargin = cm(sheetVariable(editorCssVariables.pageMarginLeft));

  /** A twip of paper is under a tenth of a pixel, which is the whole margin of error here */
  const ONE_TWIP_PX = 96 / 1440;

  it("writes down the same paper the fallback geometry names", () => {
    expect(pageWidth * PX_PER_CM).toBeCloseTo(A4_PAGE_PIXELS.pageWidth, 1);
    expect(pageHeight * PX_PER_CM).toBeCloseTo(A4_PAGE_PIXELS.pageHeight, 1);
  });

  it("writes down the same margins the fallback geometry names", () => {
    expect(
      Math.abs(margin * PX_PER_CM - A4_PAGE_PIXELS.marginTop)
    ).toBeLessThan(ONE_TWIP_PX);
    expect(
      Math.abs(sideMargin * PX_PER_CM - A4_PAGE_PIXELS.marginLeft)
    ).toBeLessThan(ONE_TWIP_PX);
  });

  it("leaves the body the fallback arithmetic works from", () => {
    // The CSS writes the paper in cm and the geometry rounds it to whole twips, so the two
    // part company by a fraction of a twip. That is under a tenth of a pixel and invisible;
    // what matters is that neither drifts a whole twip from the other
    const withinTwips = (css: number, geometry: number, twips: number) =>
      expect(Math.abs(css - geometry)).toBeLessThan(twips * ONE_TWIP_PX);

    withinTwips(
      (pageHeight - margin * 2) * PX_PER_CM,
      A4_PAGE_PIXELS.bodyHeight,
      2
    );
    withinTwips(
      (pageWidth - sideMargin * 2) * PX_PER_CM,
      A4_PAGE_PIXELS.bodyWidth,
      2
    );
    withinTwips(margin * PX_PER_CM, A4_PAGE_PIXELS.marginTop, 1);
  });
});
