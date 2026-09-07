/**
 * What the page engine knows about a block, whatever shape of block it is.
 *
 * A measurer says where a block may be parted and how much of it has to fit on the page it
 * starts on; the layout (`page/pageLayout`) answers with the places it decided to part it at.
 * Neither side names a paragraph or a table.
 *
 * Which measurer a block gets, and what draws the cuts it was given, is the block's kind
 * (`page/kinds`).
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Decoration, EditorView } from "prosemirror-view";

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

/** One block as it stands on the sheet, handed to the kind about to measure it */
export interface MeasureTarget {
  view: EditorView;
  node: PMNode;
  pos: number;
  dom: HTMLElement;
  /**
   * Sheet coordinates from viewport coordinates: the visual scale taken out, and everything the
   * engine has opened up above this block taken back off
   */
  sheetY(viewportY: number): number;
  /** The block's natural top: pushes and spaces opened above it already taken back off */
  top: number;
  /** The visual scale the sheet is drawn at, which a length read off the screen is divided by */
  scale: number;
}

/** What a kind has to say about the block it measured */
export interface KindMeasure {
  candidates: readonly BreakCandidate[];
  /** The smallest piece that has to fit on the page the block starts on */
  minFirstPiece: number;
  /** The block holds a break it could not open a space at, so the next block starts a page */
  breakAfter: boolean;
  /** Height the engine's own marks add inside this block, taken off its measured bottom */
  appliedHeight: number;
}

/**
 * Everything the engine does with one shape of breakable block: how one is measured, and how the
 * cuts the layout gave it are drawn.
 *
 * A shape the engine learns later - a footnote area, a paragraph with a floating object hanging
 * out of it - is one more kind rather than another branch in the measurer, in the decorations
 * and in the plugin state at once.
 */
export interface BlockKind {
  readonly name: string;
  matches(node: PMNode): boolean;
  measure(target: MeasureTarget): KindMeasure;
  /** Whether `at` still names a place this kind can cut at in `doc`; a false drops the cut */
  holdsCut(doc: PMNode, at: number): boolean;
  /**
   * The decorations one block of this kind carries: the cuts the layout gave it, plus whatever it
   * needs before any measurement (a paragraph puts an empty space on every page `br`, so the
   * measurement has an element to read it off).
   */
  decorate(
    pos: number,
    node: PMNode,
    cuts: readonly PageCut[],
    into: Decoration[]
  ): void;
}

/** The first kind whose `matches` answers. The kind registered last has to match every block */
export function blockKindFor(
  kinds: readonly BlockKind[],
  node: PMNode
): BlockKind {
  const found = kinds.find((kind) => kind.matches(node));
  if (!found) {
    throw new Error(`no block kind matches a ${node.type.name} block`);
  }
  return found;
}
