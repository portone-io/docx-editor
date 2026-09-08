/**
 * The arithmetic that makes one continuous sheet look like separate pages.
 *
 * The push amounts, the space opened at each page break and the places where a gap opens up
 * between pages are all decided here in one pass, and the view simply uses those values.
 * A page break is the one thing split inside a block; nothing else is broken line by line, so
 * the places where Word actually breaks may differ.
 *
 * The size of a page is the document's own (`docx/pageGeometry`), not a fixed A4: a Letter
 * document breaks at Letter's height and fits its tables to Letter's width. A document written in
 * several sections breaks each block at the height of the section that block sits in
 * (`docx/sections`), and a section opens a page of its own unless it declares `continuous`, which
 * carries it on down the page the section before it ends on.
 *
 * The width is the exception: one sheet is drawn at one width (`styles/editor.css`), which is the
 * first section's, so a section on wider paper is paginated on its own height but drawn on the
 * paper the sheet holds.
 */

import {
  A4_PORTRAIT,
  bodyHeightTwips,
  bodyWidth,
  type PageGeometry,
  twipsToPx,
} from "../docx/pageGeometry";
import type { DocumentSection, SectionProperties } from "../docx/sections";
import { editorCssVariables } from "../styles/classNames";
import type { MeasuredBlock, PageCut } from "./blockKinds";

export type { MeasuredBlock };

/** The paper measured in the pixels the sheet is drawn with */
export interface PagePixels {
  /** The whole sheet, padding included */
  pageWidth: number;
  pageHeight: number;
  /** The width one line of body text occupies: the sheet less its side padding */
  bodyWidth: number;
  /** The height the body occupies on one page */
  bodyHeight: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  /** From the end of one page's body to the top of the next page's body */
  pageStep: number;
}

/** The grey gap visible between one page and the next. The one number the paper does not decide */
export const PAGE_SPLIT_PX = 24;

/**
 * The paper as the sheet draws it.
 *
 * Every width and height on screen comes from here, so the document's own `w:pgSz` and
 * `w:pgMar` decide the sheet, the page breaks, and the width a table or an image is fitted to.
 */
export function pagePixels(geometry: PageGeometry): PagePixels {
  const marginTop = twipsToPx(geometry.marginTopTwips);
  const marginBottom = twipsToPx(geometry.marginBottomTwips);
  return {
    pageWidth: twipsToPx(geometry.widthTwips),
    pageHeight: twipsToPx(geometry.heightTwips),
    // The same width a table is fitted to, so the sheet and the tables on it cannot drift
    bodyWidth: bodyWidth(geometry).px,
    bodyHeight: twipsToPx(bodyHeightTwips(geometry)),
    marginLeft: twipsToPx(geometry.marginLeftTwips),
    marginRight: twipsToPx(geometry.marginRightTwips),
    marginTop,
    marginBottom,
    pageStep: marginTop + marginBottom + PAGE_SPLIT_PX,
  };
}

/** The paper a document that names none is drawn on */
export const A4_PAGE_PIXELS: PagePixels = pagePixels(A4_PORTRAIT);

/** The paper one section of a document is laid out on */
export interface SectionPixels {
  /**
   * The last position this section covers. A section is closed by the paragraph carrying its
   * `w:sectPr` (§17.6.17), so the block after that one opens the section that follows, and the
   * section the body itself closes reaches past every block.
   */
  untilPos: number;
  pixels: PagePixels;
  /**
   * How this section starts (§17.6.22). `continuous` carries it on down the page the section
   * before it ends on; every other kind opens a page.
   */
  type: SectionProperties["type"];
}

/** The one section a document naming none is laid out on */
export const A4_SECTION_PIXELS: readonly SectionPixels[] = [
  { untilPos: Number.POSITIVE_INFINITY, pixels: A4_PAGE_PIXELS, type: null },
];

/**
 * The paper of every section of a document, in the pixels the sheet is drawn with.
 *
 * Reading the sections walks every block (`docx/sections`), so this runs once per document change
 * rather than once per measurement.
 */
export function sectionPixels(
  sections: readonly DocumentSection[]
): SectionPixels[] {
  return sections.map(({ anchor, props }) => ({
    untilPos:
      anchor.kind === "paragraph" ? anchor.pos : Number.POSITIVE_INFINITY,
    pixels: pagePixels(props.geometry),
    type: props.type,
  }));
}

/** Which section the block at this position belongs to: the first one reaching that far */
function sectionIndexAt(
  sections: readonly SectionPixels[],
  pos: number
): number {
  const reaching = sections.findIndex((section) => pos <= section.untilPos);
  return reaching === -1 ? sections.length - 1 : reaching;
}

/**
 * The paper the block at this position is laid out on, which is the paper of every page that
 * block stands on. A4 where the list names no section at all.
 */
export function sectionPaperAt(
  sections: readonly SectionPixels[],
  pos: number
): PagePixels {
  return sections[sectionIndexAt(sections, pos)]?.pixels ?? A4_PAGE_PIXELS;
}

/** A pixel length as CSS writes it, rounded so the sheet and the arithmetic agree to the pixel */
function px(value: number): string {
  return `${Math.round(value * 100) / 100}px`;
}

/**
 * The paper the document names, as the CSS variables the sheet is drawn from.
 *
 * The page breaks above are worked out from the same geometry, so the divider drawn between
 * pages lands where the paper actually ends. `editor.css` carries A4 as the fallback for the
 * moment before a document is opened.
 */
export function pageGeometryStyle(page: PagePixels): string {
  return [
    `${editorCssVariables.pageWidth}:${px(page.pageWidth)}`,
    `${editorCssVariables.pageHeight}:${px(page.pageHeight)}`,
    `${editorCssVariables.pageMarginTop}:${px(page.marginTop)}`,
    `${editorCssVariables.pageMarginRight}:${px(page.marginRight)}`,
    `${editorCssVariables.pageMarginBottom}:${px(page.marginBottom)}`,
    `${editorCssVariables.pageMarginLeft}:${px(page.marginLeft)}`,
  ].join(";");
}

/**
 * The slack that keeps a page from flipping when the measured values wobble below the
 * decimal point
 */
const TOLERANCE_PX = 0.5;

/** One block to be moved down to the next page */
export interface BlockPush {
  pos: number;
  /** The gap to place above the block: the original gap plus the push */
  marginTop: number;
  /** Of that, the extra amount pushed for the page's sake */
  push: number;
}

/** A place where one page parts from the next */
export interface PageSplit {
  /** Where the previous page's body ends, measured from the top of the body */
  y: number;
  /** The number of the page that follows this place */
  page: number;
  /** Whether the document records a break at this place */
  forced: boolean;
  /**
   * Whether the block is taller than one page, so it cannot be pushed and the text
   * crosses this place
   */
  crossed: boolean;
}

/** Where one page starts */
export interface PageStart {
  page: number;
  /** Where this page's body starts, measured from the top of the body */
  bodyStart: number;
  /**
   * The position of the block this page opens with, which is the block the text crossing into it
   * belongs to. 0 for a document holding no block at all
   */
  pos: number;
  /**
   * Whether text crosses over from the previous page, so this page continues with no top
   * margin
   */
  crossed: boolean;
}

export interface PageLayout {
  /** Only the blocks to be moved down to the next page */
  pushes: BlockPush[];
  /** Where the layout parted a block, and the space it opened before the continued piece */
  cuts: PageCut[];
  splits: PageSplit[];
  /** One per page, the first page included */
  pages: PageStart[];
  /** The body height with the last page filled out in full */
  bodyHeight: number;
  /**
   * The margin under that last page, which is its own section's. The sheet is padded from the
   * first section's paper alone (`editor/createEditor`), so a document ending on a deeper bottom
   * margin needs the sheet drawn down to this.
   */
  marginBottom: number;
}

export interface PageLayoutInput {
  blocks: readonly MeasuredBlock[];
  /**
   * The paper of each section in document order (`sectionPixels`). A block is laid out on the
   * paper of the section it sits in, and the section a block opens starts a page of its own
   * unless that section is `continuous`.
   */
  sections: readonly SectionPixels[];
}

/** Where one block is laid out: the paper of its section, and whether it opens a page */
interface BlockPaper {
  paper: PagePixels;
  /**
   * Whether this block opens a section that asks for a page of its own, which is every kind of
   * section start but `continuous` (§17.6.22)
   */
  opensPage: boolean;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/**
 * Whether the keep a block asks for can hold between it and the block after it.
 * A block parted inside itself starts its last piece where the cut put it, so there is nothing a
 * keep could move; and a page the document starts between the two is one no keep can close, the
 * start of a section that opens one included.
 */
function keepsWithNext(
  block: MeasuredBlock,
  next: MeasuredBlock,
  opensPage: boolean
): boolean {
  return (
    block.keepWithNext &&
    !opensPage &&
    block.candidates.length === 0 &&
    !block.breakAfter &&
    !next.breakBefore
  );
}

/**
 * What a block kept with the next one brings onto the page it starts on, by block index: itself,
 * every block kept after it, and the first piece of the block the keeps end at (§17.3.1.14).
 *
 * A run of keeps no page can hold is let go as a whole, and none of its blocks is listed. Let go
 * from its head alone, a run longer than a page would leave a page short wherever its tail first
 * fit, for a keep that cannot be kept anyway.
 */
function keptExtents(
  blocks: readonly MeasuredBlock[],
  laid: readonly BlockPaper[]
): Map<number, number> {
  const fitsPage = (extent: number, index: number) => {
    const height = laid[index]?.paper.bodyHeight ?? 0;
    return extent <= height + TOLERANCE_PX;
  };
  const extents = new Map<number, number>();
  /** The run being walked up from its end, the block nearest its head last */
  let run: { index: number; extent: number }[] = [];
  const settle = () => {
    const head = run.at(-1);
    if (head && fitsPage(head.extent, head.index)) {
      for (const { index, extent } of run) {
        extents.set(index, extent);
      }
    }
    run = [];
  };
  for (let index = blocks.length - 2; index >= 0; index -= 1) {
    const block = blocks[index];
    const next = blocks[index + 1];
    const opensPage = laid[index + 1]?.opensPage === true;
    if (block && next && keepsWithNext(block, next, opensPage)) {
      const below = run.at(-1)?.extent ?? next.minFirstPiece;
      run.push({ index, extent: block.height + next.gap + below });
    } else {
      settle();
    }
  }
  settle();
  return extents;
}

/**
 * Distributes the blocks across pages.
 *
 * A block straddling the end of a page is pushed to the top of the next page, and the
 * blocks below it move down by the same amount.
 * A block taller than one page cannot be pushed, so it is left where it is and only the
 * places it crosses are reported.
 * A block is parted only where its measurer offered a candidate. A forced one - a page break
 * written into the text - is filled out to the end of the page it falls on, which carries the rest
 * of that block to the top of the next page while the block itself stays whole. An optional one -
 * a row boundary in a table - is taken only where the piece after it would run off the page.
 * A block kept with the next one is pushed together with what it is kept with: the blocks kept
 * after it and the first piece of the block the keeps end at.
 * Each block is measured against the paper of its own section, and the first block of a section
 * opens a page of its own unless the section it opens declares `continuous`.
 */
export function pageLayout({ blocks, sections }: PageLayoutInput): PageLayout {
  const pushes: BlockPush[] = [];
  const cuts: PageCut[] = [];
  const splits: PageSplit[] = [];
  /** The block being placed, which is the one a page opened along the way starts with */
  let blockPos = blocks[0]?.pos ?? 0;
  const firstPage: PageStart = {
    page: 1,
    bodyStart: 0,
    pos: blockPos,
    crossed: false,
  };
  const pages: PageStart[] = [firstPage];
  // A section is read into a geometry that always leaves a body to draw on
  // (`docx/pageGeometry`), so a paper with no body is a hand-built one; a list with no such paper
  // at all cannot hold a single page, let alone a boundary between two
  const roomy = sections.find(
    (section) => section.pixels.bodyHeight > 0
  )?.pixels;
  if (roomy === undefined) {
    return { pushes, cuts, splits, pages, bodyHeight: 0, marginBottom: 0 };
  }

  /** The section each block sits in, in document order */
  const blockSections = blocks.map((block) =>
    sectionIndexAt(sections, block.pos)
  );
  /** Where each block is laid out, read once for the whole pass */
  const laid: BlockPaper[] = blockSections.map((section, index) => {
    const pixels = sections[section]?.pixels;
    return {
      paper: pixels && pixels.bodyHeight > 0 ? pixels : roomy,
      opensPage:
        index > 0 &&
        section !== blockSections[index - 1] &&
        sections[section]?.type !== "continuous",
    };
  });

  let pageStart = 0;
  let cursor = 0;
  /** The paper of the block being placed, which is the paper of the page it stands on */
  let paper = laid[0]?.paper ?? roomy;

  /**
   * From the end of the body of the page being closed to the top of the body of the page opening:
   * the closed page keeps its own bottom margin and the opening one its own top margin, which
   * comes to the same step on both sides of a boundary inside one section.
   */
  const stepTo = (opening: PagePixels) =>
    paper.pageStep - paper.marginTop + opening.marginTop;

  const split = (
    y: number,
    crossed: boolean,
    forced: boolean,
    opening: PagePixels = paper
  ) => {
    const page = splits.length + 2;
    splits.push({ y: round(y), page, forced, crossed });
    pageStart = crossed ? y : y + stepTo(opening);
    pages.push({ page, bodyStart: round(pageStart), pos: blockPos, crossed });
  };

  // No gap is placed along a stretch the text crosses: what a block taller than one page covers,
  // and the part of a block that runs past the end of its page before the next break in it
  const crossTo = (y: number) => {
    while (y > pageStart + paper.bodyHeight + TOLERANCE_PX) {
      split(pageStart + paper.bodyHeight, true, false);
    }
  };

  const kept = keptExtents(blocks, laid);
  let breakAfterPrevious = false;

  for (const [index, block] of blocks.entries()) {
    blockPos = block.pos;
    /** The paper this block is laid out on, which is the paper of the page it opens */
    const opening = laid[index]?.paper ?? paper;
    const startsSection = laid[index]?.opensPage === true;
    // The page being closed was filled with the blocks already on it, so it ends where their own
    // paper ends rather than where this block's does
    const pageEnd = pageStart + paper.bodyHeight;
    const top = cursor + block.gap;
    const startsPage =
      (block.breakBefore || breakAfterPrevious || startsSection) &&
      top > pageStart + TOLERANCE_PX;
    // Only the piece up to the first candidate has to fit on the page the block starts on, and
    // for a block kept with the next one, whatever it is kept with as well
    const first = kept.get(index) ?? block.minFirstPiece;
    const overflows = top + first > pageEnd + TOLERANCE_PX;
    const fits = first <= opening.bodyHeight + TOLERANCE_PX;

    const push =
      startsPage || (overflows && fits)
        ? Math.max(0, pageEnd + stepTo(opening) - top)
        : 0;
    if (push > 0 || startsPage) split(pageEnd, false, startsPage, opening);
    // Whether or not a page was opened for it, this block and everything after it stand on its
    // own section's paper
    paper = opening;
    if (push > 0) {
      pushes.push({
        pos: block.pos,
        marginTop: round(block.gap + push),
        push: round(push),
      });
    }

    // Offsets are read off the block with no space in it, so each cut shifts the ones after it
    const contentTop = top + push;
    const pageBodyHeight = paper.bodyHeight;
    let added = 0;
    let pieceStart = 0;
    for (let index = 0; index <= block.candidates.length; index += 1) {
      const cut = index === 0 ? null : block.candidates[index - 1];
      const pieceEnd = block.candidates[index]?.offset ?? block.height;
      const pieceTop = contentTop + pieceStart + added;
      let pieceBottom = contentTop + pieceEnd + added;
      if (cut) {
        // A forced cut is answered from the page its own place stands on, however many pages
        // the piece before it crossed to get there
        if (cut.forced) crossTo(pieceTop);
        const pageBottom = pageStart + pageBodyHeight;
        const cutHere =
          cut.forced ||
          (pieceBottom > pageBottom + TOLERANCE_PX &&
            pieceTop > pageStart + TOLERANCE_PX &&
            pieceEnd - pieceStart <= pageBodyHeight + TOLERANCE_PX);
        if (cutHere) {
          split(pageBottom, false, cut.forced);
          const height = Math.max(0, pageStart - pieceTop);
          cuts.push({ at: cut.at, height: round(height) });
          added += height + cut.repeatHeight;
          pieceBottom += height + cut.repeatHeight;
        }
      }
      crossTo(pieceBottom);
      pieceStart = pieceEnd;
    }

    cursor = contentTop + block.height + added;
    breakAfterPrevious = block.breakAfter;
  }

  return {
    pushes,
    cuts,
    splits,
    pages,
    // The last page is filled out in full on the paper of the section it opens with
    bodyHeight: round(pageStart + paper.bodyHeight),
    marginBottom: paper.marginBottom,
  };
}
