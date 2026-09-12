/**
 * The kind every block falls back to: a paragraph, and any other shape no kind ahead of it claims.
 *
 * A page `br` is the one place a paragraph is parted at, and the space opened at it is what the
 * measurement reads the break's place off, so every `br` carries a space of its own before
 * anything is measured. A break standing where no space can be opened - inside a table cell,
 * where it would grow the cell rather than the page - is answered by the whole block instead.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Decoration } from "prosemirror-view";
import { toParagraphFormat } from "../../model/format";
import { docxSchema, isPageBreak } from "../../schema";
import { editorAttributes } from "../../styles/classNames";
import type { BlockKind, BreakCandidate } from "../blockKinds";

const PAGE_BREAK_BR = `br[${editorAttributes.breakType}="page"]`;
const BREAK_SPACE = `[${editorAttributes.pageBreakSpace}]`;

/**
 * A block box around the `br`, so what follows the break is laid out below it and a height moves
 * it down by exactly that much. An empty one redraws the paragraph as it stood, except one holding
 * nothing but the break, which loses the line the `br` alone occupied.
 */
function spaceStyle(height: number): string {
  return `display:block;height:${height}px`;
}

function isPageBreakNode(node: PMNode | null | undefined): boolean {
  return (
    node?.type === docxSchema.nodes.hardBreak && isPageBreak(node.attrs.brAttrs)
  );
}

/** One page break inside a block, where it stands and how much room the `br` itself takes */
interface BreakAt {
  at: number;
  size: number;
}

/**
 * Every page break inside this block, in document order.
 * A space is put on the breaks of a top-level paragraph alone: inside a table cell it would grow
 * the cell rather than the page, so a break there is left to the whole-block rule below, and none
 * is found here.
 *
 * The measurement reads the same list to pair each space element it finds with the break it
 * belongs to.
 */
function pageBreaksIn(block: PMNode, blockPos: number): BreakAt[] {
  const found: BreakAt[] = [];
  if (block.type !== docxSchema.nodes.paragraph) return found;
  block.forEach((child, offset) => {
    if (isPageBreakNode(child)) {
      found.push({ at: blockPos + 1 + offset, size: child.nodeSize });
    }
  });
  return found;
}

export const paragraphKind: BlockKind = {
  name: "paragraph",

  matches: () => true,

  measure({ node, pos, dom, sheetY, top, scale }) {
    const breaks = pageBreaksIn(node, pos);
    const candidates: BreakCandidate[] = [];
    /** What the spaces read so far have opened up inside this block */
    let appliedHeight = 0;
    /** What each break's own space is drawn at, which is 0 until the layout cuts there */
    const opened = new Map<number, number>();
    // Each space element is the one the break at the same ordinal was given
    dom.querySelectorAll(BREAK_SPACE).forEach((space, index) => {
      const box = space.getBoundingClientRect();
      const found = breaks.at(index);
      const drawn = box.height / scale;
      if (found) {
        candidates.push({
          at: found.at,
          offset: sheetY(box.top) - appliedHeight - top,
          forced: true,
          repeatHeight: 0,
        });
        if (drawn > 0) opened.set(found.at, drawn);
      }
      appliedHeight += drawn;
    });
    return {
      candidates,
      minFirstPiece:
        candidates[0]?.offset ??
        dom.getBoundingClientRect().height / scale - appliedHeight,
      // A break with no space of its own sits somewhere the space could not be opened. Those
      // still start a new page, but only after the whole block
      breakAfter:
        dom.querySelectorAll(PAGE_BREAK_BR).length > candidates.length,
      appliedHeight,
      opened,
      keepWithNext: toParagraphFormat(node.attrs.format)?.keepNext === true,
    };
  },

  holdsCut: (doc, at) => isPageBreakNode(doc.nodeAt(at)),

  decorate(pos, node, cuts, into) {
    const heights = new Map(cuts.map((cut) => [cut.at, cut.height]));
    for (const { at, size } of pageBreaksIn(node, pos)) {
      const height = heights.get(at) ?? 0;
      into.push(
        Decoration.inline(at, at + size, {
          nodeName: "span",
          style: spaceStyle(height),
          [editorAttributes.pageBreakSpace]: `${height}`,
        })
      );
    }
  },
};
