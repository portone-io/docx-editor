/**
 * The arithmetic that makes one continuous sheet look like separate pages.
 *
 * The push amounts, the space opened at each page break and the places where a gap opens up
 * between pages are all decided here in one pass, and the view simply uses those values.
 * A page break is the one thing split inside a block; nothing else is broken line by line, so
 * the places where Word actually breaks may differ.
 *
 * The size of a page is the document's own (`docx/pageGeometry`), not a fixed A4: a Letter
 * document breaks at Letter's height and fits its tables to Letter's width.
 */

import {
  A4_PORTRAIT,
  bodyHeightTwips,
  bodyWidth,
  type PageGeometry,
  twipsToPx,
} from "../docx/pageGeometry";
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
   * Whether every place before this one was recorded by the document, so the number can
   * be trusted
   */
  exactPage: boolean;
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
   * Whether every place before this one was recorded by the document, so the number can
   * be trusted
   */
  exactPage: boolean;
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
}

export interface PageLayoutInput {
  blocks: readonly MeasuredBlock[];
  pageBodyHeight: number;
  /** From the end of the previous page's body to the top of the next page's body */
  pageStep: number;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
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
 */
export function pageLayout({
  blocks,
  pageBodyHeight,
  pageStep,
}: PageLayoutInput): PageLayout {
  const pushes: BlockPush[] = [];
  const cuts: PageCut[] = [];
  const splits: PageSplit[] = [];
  const firstPage: PageStart = {
    page: 1,
    bodyStart: 0,
    exactPage: true,
    crossed: false,
  };
  const pages: PageStart[] = [firstPage];
  if (!(pageBodyHeight > 0)) {
    return { pushes, cuts, splits, pages, bodyHeight: 0 };
  }

  let pageStart = 0;
  let cursor = 0;
  let exactPage = true;

  const split = (y: number, crossed: boolean, forced: boolean) => {
    if (!forced) exactPage = false;
    const page = splits.length + 2;
    splits.push({ y: round(y), page, forced, exactPage, crossed });
    pageStart = crossed ? y : y + pageStep;
    pages.push({ page, bodyStart: round(pageStart), exactPage, crossed });
  };

  // No gap is placed along a stretch the text crosses: what a block taller than one page covers,
  // and the part of a block that runs past the end of its page before the next break in it
  const crossTo = (y: number) => {
    while (y > pageStart + pageBodyHeight + TOLERANCE_PX) {
      split(pageStart + pageBodyHeight, true, false);
    }
  };

  let breakAfterPrevious = false;

  for (const block of blocks) {
    const pageEnd = pageStart + pageBodyHeight;
    const top = cursor + block.gap;
    const startsPage =
      (block.breakBefore || breakAfterPrevious) &&
      top > pageStart + TOLERANCE_PX;
    // Only the piece up to the first candidate has to fit on the page the block starts on
    const first = block.minFirstPiece;
    const overflows = top + first > pageEnd + TOLERANCE_PX;
    const fits = first <= pageBodyHeight + TOLERANCE_PX;

    const push =
      startsPage || (overflows && fits)
        ? Math.max(0, pageEnd + pageStep - top)
        : 0;
    if (push > 0 || startsPage) split(pageEnd, false, startsPage);
    if (push > 0) {
      pushes.push({
        pos: block.pos,
        marginTop: round(block.gap + push),
        push: round(push),
      });
    }

    // Offsets are read off the block with no space in it, so each cut shifts the ones after it
    const contentTop = top + push;
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
    bodyHeight: round(pageStart + pageBodyHeight),
  };
}
