/**
 * What the page engine knows about a block, whatever shape of block it is.
 *
 * A measurer says where a block may be parted and how much of it has to fit on the page it
 * starts on; the layout (`page/pageLayout`) answers with the places it decided to part it at.
 * Neither side names a paragraph or a table.
 */

/** A place inside a block where the next page may, or must, start */
export interface BreakCandidate {
  /**
   * The document position the continued piece starts at: a page `br` for a paragraph, a row for a
   * table. A position maps through an edit, which is why the marks are kept by position and not
   * by ordinal (see `page/pageDecorations`)
   */
  at: number;
  /** The top of the continued piece, measured from the block's natural top with no space in it */
  offset: number;
  /** A forced candidate always cuts; an optional one cuts only when the piece after it would overflow */
  forced: boolean;
  /** Height the continued piece carries onto the next page ahead of its own content (a repeated header) */
  repeatHeight: number;
}

/** One body block as drawn on screen */
export interface MeasuredBlock {
  /** The position where this block starts in the document */
  pos: number;
  /** The height that naturally opens up between the previous block and this one */
  gap: number;
  height: number;
  /** The document records that a new page starts at this block */
  breakBefore: boolean;
  /** The block holds a break it cannot open a space at (inside a cell), so the next block starts a page */
  breakAfter: boolean;
  candidates: readonly BreakCandidate[];
  /** The smallest piece that has to fit on the page the block starts on */
  minFirstPiece: number;
}

/** A page cut the layout decided on: the space opened before the continued piece */
export interface PageCut {
  at: number;
  height: number;
}
