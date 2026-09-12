// @vitest-environment node
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { LETTER_GEOMETRY } from "../__testing__/docx";
import { DEFAULT_SECTION, type DocumentSection } from "../docx/sections";
import { editorClassNames, editorCssVariables } from "../styles/classNames";
import type { BreakCandidate } from "./blockKinds";
import type { DemandBand, PageDemand } from "./demands";
import {
  A4_PAGE_PIXELS,
  type MeasuredBlock,
  PAGE_SPLIT_PX,
  pageGeometryStyle,
  pageLayout,
  pagePixels,
  type SectionPixels,
  sectionPaperAt,
  sectionPixels,
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
      keepWithNext: false,
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

/** One section covering the whole document, on paper 1000 tall with a 200 step between pages */
const ONE_SECTION: readonly SectionPixels[] = [
  {
    untilPos: Number.POSITIVE_INFINITY,
    pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: STEP },
    type: null,
  },
];

/**
 * The same paper for the two blocks the first section covers, and paper 1600 tall for every
 * block after them: a landscape section following a portrait one
 */
const LANDSCAPE_PAGE = 1600;
const TWO_SECTIONS: readonly SectionPixels[] = [
  {
    untilPos: 10,
    pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: STEP },
    type: null,
  },
  {
    untilPos: Number.POSITIVE_INFINITY,
    pixels: { ...A4_PAGE_PIXELS, bodyHeight: LANDSCAPE_PAGE, pageStep: STEP },
    type: null,
  },
];

/**
 * The same two sections on the same paper, the second one declaring `continuous`: a section that
 * carries on down the page the one before it ends on (§17.6.22)
 */
const CONTINUOUS_SECTIONS: readonly SectionPixels[] = [
  {
    untilPos: 10,
    pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: STEP },
    type: null,
  },
  {
    untilPos: Number.POSITIVE_INFINITY,
    pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, pageStep: STEP },
    type: "continuous",
  },
];

function layout(list: MeasuredBlock[]) {
  return pageLayout({ blocks: list, sections: ONE_SECTION });
}

describe("pageLayout", () => {
  it("nothing to push and nothing to split when it all fits on one page", () => {
    const result = layout(blocks(300, 300, 300));
    expect(result.pushes).toEqual([]);
    expect(result.splits).toEqual([]);
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, pos: 0, pageInSection: 1, crossed: false },
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

  it("reports where each page starts and which block it opens with", () => {
    const result = layout(blocks(900, 300));
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, pos: 0, pageInSection: 1, crossed: false },
      {
        page: 2,
        bodyStart: PAGE + STEP,
        pos: 10,
        pageInSection: 2,
        crossed: false,
      },
    ]);
  });

  it("counts each page again from one at the top of the section it opens", () => {
    // The third block opens the landscape section and is taller than its 1600 of body, so it
    // crosses from that section's first page onto its second
    const result = pageLayout({
      blocks: blocks(300, 300, 2000),
      sections: TWO_SECTIONS,
    });

    expect(result.pages.map((page) => [page.pos, page.pageInSection])).toEqual([
      [0, 1],
      [20, 1],
      [20, 2],
    ]);
  });

  it("a continuous section does not restart the count on the page it shares", () => {
    // The third block opens the continuous section partway down page 2, so that page stays with
    // the section that opened it and only the page after it starts the new section's count
    const result = pageLayout({
      blocks: blocks(900, 300, 300, 900),
      sections: CONTINUOUS_SECTIONS,
    });

    expect(result.pages.map((page) => [page.pos, page.pageInSection])).toEqual([
      [0, 1],
      [10, 2],
      [30, 1],
    ]);
  });

  it("counts every page of a document written in one section", () => {
    expect(
      layout(blocks(900, 300, 900)).pages.map((page) => page.pageInSection)
    ).toEqual([1, 2, 3]);
  });

  it("names the block a page continues, not the one after it", () => {
    // The one block is two pages and a half tall, so both pages after the first continue it
    const result = layout(blocks(300, 2500));
    expect(result.pages.map((page) => page.pos)).toEqual([0, 10, 10]);
  });

  it("a page reached by crossing joins onto the previous page with no top margin", () => {
    const result = layout(blocks(1500));
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, pos: 0, pageInSection: 1, crossed: false },
      { page: 2, bodyStart: PAGE, pos: 0, pageInSection: 2, crossed: true },
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

  it("a block kept with the next one is pushed together with its follower's first piece", () => {
    // The 50 alone would fit in the 100 left on the page; it goes because its follower does not
    const result = layout(blocks(900, { height: 50, keepWithNext: true }, 300));

    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 900, push: PAGE + STEP - 900 },
    ]);
    expect(result.pages).toHaveLength(2);
  });

  it("the gap above the follower is part of what is kept together", () => {
    // The 50 fits in the 100 left, and so would the 40 after it; the 20 between them is what does not
    const result = layout(
      blocks(900, { height: 50, keepWithNext: true }, { height: 40, gap: 20 })
    );

    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
  });

  it("a keep needs only the follower's first piece on the page", () => {
    // The 600 after the 50 would not fit, but only its 100 up to the break has to
    const result = layout(
      blocks(800, { height: 50, keepWithNext: true }, broken(600, 100))
    );

    expect(result.pushes).toEqual([]);
    expect(result.cuts).toEqual([{ at: 100, height: PAGE - 950 + STEP }]);
  });

  it("every block of a run of keeps moves with the head", () => {
    // 200, then 30 of gap and 150, then the 30 the keeps end at: 410 into the 400 left
    const result = layout(
      blocks(
        600,
        { height: 200, keepWithNext: true },
        { height: 150, gap: 30, keepWithNext: true },
        30
      )
    );

    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 600, push: PAGE + STEP - 600 },
    ]);
    expect(result.pages).toHaveLength(2);
  });

  it("a keep chain taller than a page is laid out block by block", () => {
    const result = layout(
      blocks(980, { height: 50, keepWithNext: true }, PAGE)
    );

    // The 50 is pushed on its own, as it would be with no keep on it, and the page after it
    expect(result.pushes.map((push) => push.pos)).toEqual([10, 20]);
    expect(result.splits.map((split) => split.crossed)).toEqual([false, false]);
  });

  it("a run of keeps no page can hold is let go as a whole, not from its head alone", () => {
    // The whole run would be 1150; the 150 and the 700 alone would be 850 and fit
    const result = layout(
      blocks(
        500,
        { height: 300, keepWithNext: true },
        { height: 150, keepWithNext: true },
        700
      )
    );

    expect(result.pushes.map((push) => push.pos)).toEqual([30]);
  });

  it("a keep on the last block changes nothing", () => {
    expect(layout(blocks(900, { height: 300, keepWithNext: true }))).toEqual(
      layout(blocks(900, 300))
    );
  });

  it("a block parted by a page break is not kept with the block after it", () => {
    // Its last piece starts the page the break opened, and the 600 follows it there
    const result = layout(
      blocks(500, { ...broken(300, 100), keepWithNext: true }, 600)
    );

    expect(result.pushes).toEqual([]);
    expect(result.cuts).toEqual([{ at: 100, height: PAGE - 600 + STEP }]);
  });

  it("a keep does not reach across a page the document starts between the two", () => {
    const before = layout(
      blocks(
        900,
        { height: 50, keepWithNext: true },
        { height: 300, breakBefore: true }
      )
    );
    expect(before.pushes.map((push) => push.pos)).toEqual([20]);
    expect(before.pages).toHaveLength(2);

    const after = layout(
      blocks(900, { height: 50, keepWithNext: true, breakAfter: true }, 300)
    );
    expect(after.pushes.map((push) => push.pos)).toEqual([20]);
    expect(after.pages).toHaveLength(2);
  });

  it("a landscape second section gets its own body height from its first block", () => {
    // 300 + 300 fill the first section's page; the 1500 that follows opens the landscape
    // section, whose 1600 of body holds it whole where the first section's 1000 could not
    const result = pageLayout({
      blocks: blocks(300, 300, 1500),
      sections: TWO_SECTIONS,
    });

    expect(result.pushes).toEqual([
      { pos: 20, marginTop: PAGE + STEP - 600, push: PAGE + STEP - 600 },
    ]);
    expect(result.splits).toEqual([
      { y: PAGE, page: 2, forced: true, crossed: false },
    ]);
    expect(result.pages).toEqual([
      { page: 1, bodyStart: 0, pos: 0, pageInSection: 1, crossed: false },
      {
        page: 2,
        bodyStart: PAGE + STEP,
        pos: 20,
        pageInSection: 1,
        crossed: false,
      },
    ]);
    // The last page is filled out on the paper it opens with, which is the landscape one
    expect(result.bodyHeight).toBe(PAGE + STEP + LANDSCAPE_PAGE);
  });

  it("a section starts a new page even where its first block would have fitted", () => {
    const result = pageLayout({
      blocks: blocks(300, 300, 100),
      sections: TWO_SECTIONS,
    });

    expect(result.pushes.map((push) => push.pos)).toEqual([20]);
    expect(result.pages).toHaveLength(2);
    expect(result.splits[0]?.forced).toBe(true);
  });

  it("a boundary between two sections keeps the margin of each page at it", () => {
    // The page that ends keeps its own bottom margin and the page that opens its own top margin,
    // so the step across the boundary is neither section's own step
    const margins: readonly SectionPixels[] = [
      {
        untilPos: 0,
        pixels: {
          ...A4_PAGE_PIXELS,
          bodyHeight: PAGE,
          pageStep: STEP,
          marginTop: 20,
        },
        type: null,
      },
      {
        untilPos: Number.POSITIVE_INFINITY,
        pixels: {
          ...A4_PAGE_PIXELS,
          bodyHeight: LANDSCAPE_PAGE,
          pageStep: 999,
          marginTop: 60,
        },
        type: null,
      },
    ];
    const across = PAGE + STEP - 20 + 60;
    const result = pageLayout({ blocks: blocks(300, 300), sections: margins });

    expect(result.pages[1]?.bodyStart).toBe(across);
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: across - 300, push: across - 300 },
    ]);
  });

  it("a continuous section carries on down the page instead of opening one", () => {
    // The same blocks that open a page at the boundary when the section starts on a new page
    const result = pageLayout({
      blocks: blocks(300, 300, 100),
      sections: CONTINUOUS_SECTIONS,
    });

    expect(result.pushes).toEqual([]);
    expect(result.splits).toEqual([]);
    expect(result.pages).toHaveLength(1);
  });

  it("a keep reaches across a continuous boundary, which opens no page to close", () => {
    const result = pageLayout({
      blocks: blocks(900, { height: 50, keepWithNext: true }, 100),
      sections: CONTINUOUS_SECTIONS,
    });

    // The kept block goes down with what it is kept with, rather than the boundary pushing the
    // block after it on its own
    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
    expect(result.pages).toHaveLength(2);
  });

  it("leaves the sheet the bottom margin of the page the document ends on", () => {
    const deeper: readonly SectionPixels[] = [
      {
        untilPos: 10,
        pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, marginBottom: 10 },
        type: null,
      },
      {
        untilPos: Number.POSITIVE_INFINITY,
        pixels: { ...A4_PAGE_PIXELS, bodyHeight: PAGE, marginBottom: 90 },
        type: null,
      },
    ];

    expect(
      pageLayout({ blocks: blocks(300, 300, 100), sections: deeper })
    ).toMatchObject({ marginBottom: 90 });
    // A document ending inside the first section is left that section's own margin
    expect(
      pageLayout({ blocks: blocks(300, 300), sections: deeper })
    ).toMatchObject({ marginBottom: 10 });
  });

  it("a keep does not reach across a section boundary", () => {
    // Without the boundary the 50 would be pushed to carry the 100 along with it
    const result = pageLayout({
      blocks: blocks(900, { height: 50, keepWithNext: true }, 100),
      sections: TWO_SECTIONS,
    });

    expect(result.pushes.map((push) => push.pos)).toEqual([20]);
    expect(result.pages).toHaveLength(2);
  });

  it("does nothing for a page height that cannot be measured", () => {
    const flat: SectionPixels[] = [
      {
        untilPos: Number.POSITIVE_INFINITY,
        pixels: { ...A4_PAGE_PIXELS, bodyHeight: 0 },
        type: null,
      },
    ];
    expect(pageLayout({ blocks: blocks(500), sections: flat })).toMatchObject({
      pushes: [],
      splits: [],
    });
    expect(pageLayout({ blocks: blocks(500), sections: [] })).toMatchObject({
      pushes: [],
      splits: [],
    });
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

describe("the room a page keeps at its foot", () => {
  const BAND = "test";

  /** The one band these demands are kept in, holding each id at the height given */
  function band(
    heights: Readonly<Record<string, number>>,
    overhead = 0
  ): ReadonlyMap<string, DemandBand> {
    return new Map([
      [BAND, { order: 0, overhead, heights: new Map(Object.entries(heights)) }],
    ]);
  }

  function demand(id: string, offset: number): PageDemand {
    return { offset, id, band: BAND };
  }

  function laidOut(
    list: MeasuredBlock[],
    bands: ReadonlyMap<string, DemandBand>
  ) {
    return pageLayout({ blocks: list, sections: ONE_SECTION, bands });
  }

  it("lays every page out as before when no band is given", () => {
    const shapes = [
      blocks(900, 300, 700),
      blocks(900, broken(300, 50)),
      blocks(broken(400, 100, 250)),
      blocks(2500, 200),
      blocks({
        height: 1500,
        candidates: rowBoundaries(100, 300, 600, 900, 1200),
        minFirstPiece: 300,
      }),
      blocks(600, { height: 200, keepWithNext: true }, 150, 30),
    ];
    for (const plain of shapes) {
      const asking = plain.map((block) => ({
        ...block,
        demands: [demand("a", 10), demand("b", block.height / 2)],
      }));
      const unchanged = layout(plain);
      expect(pageLayout({ blocks: asking, sections: ONE_SECTION })).toEqual(
        unchanged
      );
      expect(
        pageLayout({ blocks: asking, sections: ONE_SECTION, bands: new Map() })
      ).toEqual(unchanged);
      expect(unchanged.reserved).toEqual([]);
    }
  });

  it("ends a page above the room its blocks demand", () => {
    const result = laidOut(
      blocks({ height: 400, demands: [demand("a", 100)] }, 400),
      band({ a: 300 })
    );

    // The 300 kept for the first block leaves 600 of body, which the second 400 runs past
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 400, push: PAGE + STEP - 400 },
    ]);
    // The page itself still ends where its paper does
    expect(result.splits).toEqual([
      { y: PAGE, page: 2, forced: false, crossed: false },
    ]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["a"], top: PAGE - 300, height: 300 },
    ]);
  });

  it("pushes a block to the next page when its demand does not fit beneath it", () => {
    const result = laidOut(
      blocks(600, { height: 300, demands: [demand("a", 200)] }),
      band({ a: 200 })
    );

    // The 300 fits in the 400 left, but not with the 200 its own place asks for beneath it
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 600, push: PAGE + STEP - 600 },
    ]);
    expect(result.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["a"],
        top: 2 * PAGE + STEP - 200,
        height: 200,
      },
    ]);
  });

  it("counts one id demanded twice on a page once", () => {
    const result = laidOut(
      blocks(
        { height: 300, demands: [demand("a", 10), demand("a", 200)] },
        { height: 300, demands: [demand("a", 100)] }
      ),
      band({ a: 350 }, 20)
    );

    // Kept once, the 370 fits under the 600 of text; kept for every place, it would not
    expect(result.pushes).toEqual([]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["a"], top: PAGE - 370, height: 370 },
    ]);
  });

  it("ends the page a tall block crosses above that page's demands", () => {
    const result = laidOut(
      blocks({ height: 2500, demands: [demand("a", 500)] }),
      band({ a: 100 })
    );

    expect(result.splits).toEqual([
      { y: PAGE - 100, page: 2, forced: false, crossed: true },
      { y: 2 * PAGE - 100, page: 3, forced: false, crossed: true },
    ]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["a"], top: PAGE - 100, height: 100 },
    ]);
    expect(result.bodyHeight).toBe(3 * PAGE - 100);
  });

  it("reserves room on the page where a table row's demand lands", () => {
    const result = laidOut(
      blocks({
        height: 1200,
        candidates: rowBoundaries(0, 300, 600, 900),
        minFirstPiece: 300,
        demands: [demand("first row", 100), demand("third row", 650)],
      }),
      band({ "first row": 50, "third row": 150 })
    );

    // The third row ends at 900, above the 950 the first row's room leaves, but not with the 150
    // its own place asks for; with no demand the table would continue at the fourth row instead
    expect(result.cuts).toEqual([{ at: 102, height: PAGE + STEP - 600 }]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["first row"], top: PAGE - 50, height: 50 },
      {
        page: 2,
        band: BAND,
        ids: ["third row"],
        top: 2 * PAGE + STEP - 150,
        height: 150,
      },
    ]);
  });

  it("carries a demand that does not fit to the top of the next page", () => {
    const result = laidOut(
      blocks(
        { height: 900, demands: [demand("carried", 100)] },
        { height: 300, demands: [demand("own", 50)] }
      ),
      band({ carried: 200, own: 100 })
    );

    // The 900 opens its page and cannot move, and the 200 does not fit beneath it, so the next
    // page keeps it ahead of the demand of its own
    expect(result.pushes.map((push) => push.pos)).toEqual([10]);
    expect(result.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["carried", "own"],
        top: 2 * PAGE + STEP - 300,
        height: 300,
      },
    ]);
  });

  it("carries the demands a page cannot keep in the order their places stand", () => {
    const heights = band({ x: 150, y: 150 });
    const asking = {
      height: 900,
      demands: [demand("x", 100), demand("y", 200)],
    };

    // The first page lets go of the last it took first, and the next page keeps both as they stand
    const once = laidOut(blocks(asking, 300), heights);
    expect(once.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["x", "y"],
        top: 2 * PAGE + STEP - 300,
        height: 300,
      },
    ]);

    // Under a 900 it cannot move either, the second page lets both go again, and the third keeps
    // them in the same order
    const twice = laidOut(blocks(asking, 900, 50), heights);
    expect(twice.pages).toHaveLength(3);
    expect(twice.reserved).toEqual([
      {
        page: 3,
        band: BAND,
        ids: ["x", "y"],
        top: 3 * PAGE + 2 * STEP - 300,
        height: 300,
      },
    ]);
  });

  it("carries every demand that follows one its page let go", () => {
    const result = laidOut(
      blocks(
        { height: 900, demands: [demand("x", 100)] },
        { height: 50, demands: [demand("y", 10)] }
      ),
      band({ x: 200, y: 40 })
    );

    // The 50 and its 40 would fit beneath the 900, but kept there y would stand a page ahead of the
    // demand before it
    expect(result.pushes).toEqual([]);
    expect(result.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["x", "y"],
        top: 2 * PAGE + STEP - 240,
        height: 240,
      },
    ]);
  });

  it("does not push a block off the page it opens to make room for a carried demand", () => {
    const result = laidOut(
      blocks({ ...broken(900, 900), demands: [demand("a", 100)] }, 900),
      band({ a: 200 })
    );

    // The break opens the second page with the carried 200 on it, and the 900 that follows starts
    // that page, so it stays and the 200 is carried on again rather than left alone on a page
    expect(result.pushes).toEqual([]);
    expect(result.cuts).toEqual([{ at: 100, height: PAGE + STEP - 900 }]);
    expect(result.pages).toHaveLength(3);
    expect(result.reserved).toEqual([
      {
        page: 3,
        band: BAND,
        ids: ["a"],
        top: 3 * PAGE + 2 * STEP - 200,
        height: 200,
      },
    ]);
  });

  it("leaves a block where it is when an empty page could not hold it with its demand either", () => {
    const result = laidOut(
      blocks(300, { height: 600, demands: [demand("a", 100)] }),
      band({ a: 500 })
    );

    // Pushed, the 600 and its 500 would not fit the next page either, and this page's 700 would
    // be left empty for nothing
    expect(result.pushes).toEqual([]);
    expect(result.splits).toEqual([
      { y: PAGE, page: 2, forced: false, crossed: false },
    ]);
    expect(result.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["a"],
        top: 2 * PAGE + STEP - 500,
        height: 500,
      },
    ]);
  });

  it("crosses a page its room fills at the end of its paper rather than opening an empty one", () => {
    const result = laidOut(
      blocks({ height: 2500, demands: [demand("a", 0)] }),
      band({ a: PAGE })
    );

    expect(result.splits).toEqual([
      { y: PAGE, page: 2, forced: false, crossed: true },
      { y: 2 * PAGE, page: 3, forced: false, crossed: true },
    ]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["a"], top: 0, height: PAGE },
    ]);
  });

  it("opens a page past the last block for a demand carried beyond it", () => {
    const result = laidOut(
      blocks({ height: 900, demands: [demand("a", 100)] }),
      band({ a: 200 })
    );

    expect(result.splits).toEqual([
      { y: PAGE, page: 2, forced: false, crossed: false },
    ]);
    expect(result.pages).toHaveLength(2);
    expect(result.bodyHeight).toBe(PAGE + STEP + PAGE);
    expect(result.reserved).toEqual([
      {
        page: 2,
        band: BAND,
        ids: ["a"],
        top: 2 * PAGE + STEP - 200,
        height: 200,
      },
    ]);
  });

  it("keeps a demand taller than an empty page on the page it lands on", () => {
    const result = laidOut(
      blocks({ height: 100, demands: [demand("tall", 50)] }, 300),
      band({ tall: 1500 })
    );

    // No page could keep it whole, so it stays with its place, and that page's body ends with the
    // text that has to stay there, which is where its band starts
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["tall"], top: 100, height: 1500 },
    ]);
    expect(result.pushes).toEqual([
      { pos: 10, marginTop: PAGE + STEP - 100, push: PAGE + STEP - 100 },
    ]);
    expect(result.pages).toHaveLength(2);
  });

  it("takes no notice of a demand whose band is not given", () => {
    const result = laidOut(
      blocks({
        height: 2500,
        demands: [demand("a", 500), { offset: 950, id: "b", band: "other" }],
      }),
      band({ a: 100 })
    );

    // The unknown place stands below the body end the known one leaves, and holds nothing there
    expect(result.splits.map((split) => split.y)).toEqual([
      PAGE - 100,
      2 * PAGE - 100,
    ]);
    expect(result.reserved).toEqual([
      { page: 1, band: BAND, ids: ["a"], top: PAGE - 100, height: 100 },
    ]);
  });

  it("stacks a page's bands where their definitions put them, not where its text met them", () => {
    // "floor" belongs at the foot of the body and "shelf" above it, whichever of the two the
    // page's text reaches first
    const bands: ReadonlyMap<string, DemandBand> = new Map([
      ["floor", { order: 0, overhead: 0, heights: new Map([["f", 100]]) }],
      ["shelf", { order: 1, overhead: 0, heights: new Map([["s", 50]]) }],
    ]);
    const met = (first: string, second: string) =>
      pageLayout({
        blocks: blocks({
          height: 300,
          demands: [
            { offset: 10, id: first[0] ?? "", band: first },
            { offset: 20, id: second[0] ?? "", band: second },
          ],
        }),
        sections: ONE_SECTION,
        bands,
      }).reserved;

    const stacked = [
      { page: 1, band: "shelf", ids: ["s"], top: PAGE - 150, height: 50 },
      { page: 1, band: "floor", ids: ["f"], top: PAGE - 100, height: 100 },
    ];
    expect(met("floor", "shelf")).toEqual(stacked);
    expect(met("shelf", "floor")).toEqual(stacked);
  });

  it("adds a band's overhead once to each page that holds it", () => {
    const bands: ReadonlyMap<string, DemandBand> = new Map([
      [
        "wide",
        {
          order: 1,
          overhead: 30,
          heights: new Map([
            ["w1", 100],
            ["w2", 50],
            ["w3", 70],
          ]),
        },
      ],
      ["narrow", { order: 0, overhead: 10, heights: new Map([["n1", 20]]) }],
    ]);
    const result = pageLayout({
      blocks: blocks(
        {
          height: 300,
          demands: [
            { offset: 10, id: "w1", band: "wide" },
            { offset: 20, id: "n1", band: "narrow" },
          ],
        },
        { height: 480, demands: [{ offset: 50, id: "w2", band: "wide" }] },
        { height: 600, demands: [{ offset: 100, id: "w3", band: "wide" }] }
      ),
      sections: ONE_SECTION,
      bands,
    });

    // The 480 and the 50 it asks for end at 830, above the 840 the first block's room leaves; with
    // the wide band's 30 counted again they would not
    expect(result.pushes.map((push) => push.pos)).toEqual([20]);
    // A page's bands stand one under the next, in the order their definitions name
    expect(result.reserved).toEqual([
      {
        page: 1,
        band: "wide",
        ids: ["w1", "w2"],
        top: PAGE - 30 - 180,
        height: 30 + 100 + 50,
      },
      {
        page: 1,
        band: "narrow",
        ids: ["n1"],
        top: PAGE - 30,
        height: 10 + 20,
      },
      {
        page: 2,
        band: "wide",
        ids: ["w3"],
        top: 2 * PAGE + STEP - 100,
        height: 30 + 70,
      },
    ]);
  });
});

describe("the paper of each section", () => {
  const first: DocumentSection = {
    index: 0,
    firstBlock: 0,
    lastBlock: 1,
    anchor: { kind: "paragraph", pos: 20 },
    props: DEFAULT_SECTION,
  };
  const last: DocumentSection = {
    index: 1,
    firstBlock: 2,
    lastBlock: 3,
    anchor: { kind: "body" },
    props: { ...DEFAULT_SECTION, geometry: LETTER_GEOMETRY },
  };

  it("is the paper that section names, up to the last position it covers", () => {
    const papers = sectionPixels([first, last]);
    expect(papers.map((paper) => paper.untilPos)).toEqual([
      20,
      Number.POSITIVE_INFINITY,
    ]);
    expect(papers[0]?.pixels).toEqual(A4_PAGE_PIXELS);
    expect(papers[1]?.pixels).toEqual(pagePixels(LETTER_GEOMETRY));
  });

  it("carries the kind of start each section declares", () => {
    const carried = sectionPixels([
      { ...first, props: { ...DEFAULT_SECTION, type: "continuous" } },
      last,
    ]);
    expect(carried.map((paper) => paper.type)).toEqual(["continuous", null]);
  });

  it("lays a block out on the first section that reaches it", () => {
    const papers = sectionPixels([first, last]);
    // The paragraph carrying a break closes its own section, so it is laid out on that paper
    expect(sectionPaperAt(papers, 20)).toEqual(A4_PAGE_PIXELS);
    expect(sectionPaperAt(papers, 21)).toEqual(pagePixels(LETTER_GEOMETRY));
    // A block past every section named, and a document naming none at all, fall back
    expect(sectionPaperAt(papers.slice(0, 1), 9999)).toEqual(A4_PAGE_PIXELS);
    expect(sectionPaperAt([], 0)).toEqual(A4_PAGE_PIXELS);
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
