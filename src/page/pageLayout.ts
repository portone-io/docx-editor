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

/** One body block as drawn on screen */
export interface MeasuredBlock {
  /** The position where this block starts in the document */
  pos: number;
  /** The height that naturally opens up between the previous block and this one */
  gap: number;
  height: number;
  /** Set when the document records that a new page starts at this block */
  breakBefore: boolean;
  /**
   * Every page break inside the block, in document order, each measured from the block's own top
   * so that the block can be laid out wherever it lands.
   * That top is the one the block would be drawn at with no space in it, so a space opened at one
   * break shifts the ones after it.
   */
  breaks: readonly number[];
  /** Row boundaries available when this block is an editable table */
  table?: MeasuredTable;
}

/** A row that may start the continued part of a table */
export interface TableBoundary {
  /** The row's document position */
  pos: number;
  /** Its top measured from the table's natural top */
  offset: number;
}

/** The measurements needed to continue a table without changing the document */
export interface MeasuredTable {
  boundaries: readonly TableBoundary[];
  /** The smallest useful first piece: headers followed by one body row group */
  firstPageMinimum: number;
  repeatHeaderHeight: number;
  /** Document positions of the contiguous header rows at the start of the table */
  headerRows: readonly number[];
  /** Changes when any projected header content or formatting changes */
  headerSignature: string;
  columns: number;
}

/** The space opened up at one page break, so that what follows it starts the next page */
export interface BreakSpace {
  /** The position where the block holding the break starts */
  pos: number;
  /** Which break inside that block, counted in document order */
  index: number;
  height: number;
}

/** One block to be moved down to the next page */
export interface BlockPush {
  pos: number;
  /** The gap to place above the block: the original gap plus the push */
  marginTop: number;
  /** Of that, the extra amount pushed for the page's sake */
  push: number;
}

/** A display-only row inserted before the row that continues on the next page */
export interface TableContinuation {
  /** The position of the first original row on the next page */
  pos: number;
  /** Space from the previous page's row boundary to the next page body */
  height: number;
  /** Header rows to project immediately after the space */
  headerRows: readonly number[];
  headerSignature: string;
  columns: number;
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
  spaces: BreakSpace[];
  tableContinuations: TableContinuation[];
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
 * A page break inside a block is filled out to the end of the page it falls on, which carries
 * the rest of that block - the rest of the sentence, the rest of the list item - to the top of
 * the next page while the block itself stays whole.
 */
export function pageLayout({
  blocks,
  pageBodyHeight,
  pageStep,
}: PageLayoutInput): PageLayout {
  const pushes: BlockPush[] = [];
  const spaces: BreakSpace[] = [];
  const tableContinuations: TableContinuation[] = [];
  const splits: PageSplit[] = [];
  const firstPage: PageStart = {
    page: 1,
    bodyStart: 0,
    exactPage: true,
    crossed: false,
  };
  const pages: PageStart[] = [firstPage];
  if (!(pageBodyHeight > 0)) {
    return {
      pushes,
      spaces,
      tableContinuations,
      splits,
      pages,
      bodyHeight: 0,
    };
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

  for (const block of blocks) {
    const pageEnd = pageStart + pageBodyHeight;
    const top = cursor + block.gap;
    const startsPage = block.breakBefore && top > pageStart + TOLERANCE_PX;
    // Only the part up to the first break has to fit on the page the block starts on
    const first =
      block.breaks[0] ?? block.table?.firstPageMinimum ?? block.height;
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

    // Offsets are read off the block with no space in it, so each space shifts the ones after it
    let contentTop = top + push;
    for (const [index, offset] of block.breaks.entries()) {
      const breakY = contentTop + offset;
      crossTo(breakY);
      split(pageStart + pageBodyHeight, false, true);
      const height = Math.max(0, pageStart - breakY);
      spaces.push({ pos: block.pos, index, height: round(height) });
      contentTop += height;
    }

    let tableAdded = 0;
    if (block.table) {
      let segmentStart = 0;
      for (let index = 0; index <= block.table.boundaries.length; index += 1) {
        const segmentEnd =
          block.table.boundaries[index]?.offset ?? block.height;
        const segmentTop = contentTop + segmentStart + tableAdded;
        let segmentBottom = contentTop + segmentEnd + tableAdded;
        const pageEnd = pageStart + pageBodyHeight;
        const overflows = segmentBottom > pageEnd + TOLERANCE_PX;
        const boundary = index === 0 ? null : block.table.boundaries[index - 1];
        const segmentHeight = segmentEnd - segmentStart;

        if (
          overflows &&
          boundary &&
          segmentTop > pageStart + TOLERANCE_PX &&
          segmentHeight <= pageBodyHeight + TOLERANCE_PX
        ) {
          split(pageEnd, false, false);
          const height = Math.max(0, pageStart - segmentTop);
          tableContinuations.push({
            pos: boundary.pos,
            height: round(height),
            headerRows: block.table.headerRows,
            headerSignature: block.table.headerSignature,
            columns: block.table.columns,
          });
          tableAdded += height + block.table.repeatHeaderHeight;
          segmentBottom = contentTop + segmentEnd + tableAdded;
        }

        crossTo(segmentBottom);
        segmentStart = segmentEnd;
      }
    }

    const bottom = contentTop + block.height + tableAdded;
    crossTo(bottom);
    cursor = bottom;
  }

  return {
    pushes,
    spaces,
    tableContinuations,
    splits,
    pages,
    bodyHeight: round(pageStart + pageBodyHeight),
  };
}
