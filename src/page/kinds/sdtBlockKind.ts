/**
 * The kind of a block-level content control: it is parted wherever the blocks it holds are.
 *
 * The container draws no box of its own - no padding, no border, no margin, and what a shut
 * control wears is an outline and a background, which take no room (`styles/editor.css`) - so the
 * control stands exactly as tall as the blocks inside it and a child's offset within it needs
 * nothing subtracted for chrome.
 *
 * Nothing therefore has to be drawn at a cut. The outline is one box around everything the control
 * holds, so it runs down both pieces and closes under the last of them, and only the background a
 * shut control paints has to be broken, which the space opened at the cut does.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Decoration, type EditorView } from "prosemirror-view";
import { isBlockControl } from "../../schema/controlAttrs";
import { editorAttributes } from "../../styles/classNames";
import {
  type BlockKind,
  type BreakCandidate,
  blockKindFor,
  type KindMeasure,
  opensPage,
  type PageCut,
} from "../blockKinds";

const CONTAINER_SPACE = editorAttributes.containerPageSpace;

/** One block of a control as the container hands it to its own kind */
interface Held {
  node: PMNode;
  pos: number;
  dom: HTMLElement;
  index: number;
}

/**
 * The spaces this very engine opened just above a block, as the browser drew them.
 * Read off the siblings rather than off a query over the control, so a nested control's spaces
 * stay that control's own and nothing walks the subtree twice.
 */
function spaceAbove(dom: HTMLElement, scale: number): number {
  let total = 0;
  let sibling = dom.previousElementSibling;
  while (
    sibling instanceof HTMLElement &&
    sibling.hasAttribute(CONTAINER_SPACE)
  ) {
    total += sibling.getBoundingClientRect().height / scale;
    sibling = sibling.previousElementSibling;
  }
  return total;
}

function heldBlocks(
  view: EditorView,
  node: PMNode,
  pos: number
): readonly Held[] {
  const found: Held[] = [];
  node.forEach((child, offset, index) => {
    const dom = view.nodeDOM(pos + 1 + offset);
    if (dom instanceof HTMLElement) {
      found.push({ node: child, pos: pos + 1 + offset, dom, index });
    }
  });
  return found;
}

/** The space a cut opens between two blocks of a control: a box as tall as the cut asked for */
function containerSpace(height: number): HTMLElement {
  const space = document.createElement("div");
  space.setAttribute(CONTAINER_SPACE, `${height}`);
  space.setAttribute("aria-hidden", "true");
  space.setAttribute("contenteditable", "false");
  space.style.height = `${height}px`;
  return space;
}

/**
 * Whether this control still holds the place a cut names. A control inside a control is walked
 * here rather than asked, since its own answer starts from the top-level block again.
 */
function heldCut(
  container: PMNode,
  containerPos: number,
  doc: PMNode,
  at: number,
  kinds: readonly BlockKind[]
): boolean {
  let holds = false;
  container.forEach((child, offset, index) => {
    if (holds) return;
    const childPos = containerPos + 1 + offset;
    if (at === childPos) holds = index > 0;
    else if (at > childPos && at < childPos + child.nodeSize) {
      holds = isBlockControl(child)
        ? heldCut(child, childPos, doc, at, kinds)
        : blockKindFor(kinds, child).holdsCut(doc, at, kinds);
    }
  });
  return holds;
}

export const sdtBlockKind: BlockKind = {
  name: "sdtBlock",

  matches: isBlockControl,

  measure({ view, node, pos, dom, kinds, sheetY, top, scale }): KindMeasure {
    const candidates: BreakCandidate[] = [];
    const opened = new Map<number, number>();
    let appliedHeight = 0;
    let minFirstPiece = dom.getBoundingClientRect().height / scale;
    let keepWithNext = false;
    let breakBefore = false;
    /** A block whose own break could not open a space asks for a page after itself */
    let breakAfter = false;

    for (const held of heldBlocks(view, node, pos)) {
      const space = spaceAbove(held.dom, scale);
      appliedHeight += space;
      const above = appliedHeight;
      const heldY = (viewportY: number) => sheetY(viewportY) - above;
      const heldTop = heldY(held.dom.getBoundingClientRect().top);
      const measured = blockKindFor(kinds, held.node).measure({
        view,
        node: held.node,
        pos: held.pos,
        dom: held.dom,
        kinds,
        sheetY: heldY,
        top: heldTop,
        scale,
      });
      const forced = breakAfter || opensPage(held.dom, measured);

      if (held.index === 0) {
        breakBefore = forced;
        // What the control has to fit on the page it starts on is what its first block does: a
        // held table still asks for its repeated headers and one body row (§17.4.78)
        minFirstPiece = heldTop - top + measured.minFirstPiece;
      } else {
        candidates.push({
          at: held.pos,
          offset: heldTop - top,
          forced,
          // A keep the block above asks for (§17.3.1.14) closes the boundary under it unless no
          // page can hold what is kept, which is the layout's to decide (`page/pageLayout`)
          ...(!forced && keepWithNext ? { kept: true } : {}),
          repeatHeight: 0,
        });
        if (space > 0) opened.set(held.pos, space);
      }

      for (const candidate of measured.candidates) {
        candidates.push({
          ...candidate,
          offset: heldTop - top + candidate.offset,
        });
      }
      for (const [at, height] of measured.opened ?? []) opened.set(at, height);
      appliedHeight += measured.appliedHeight;
      // A break the block below could not open a space at cuts at the next block instead, and
      // only a break under the last block is left for the sheet to answer after the control
      breakAfter = measured.breakAfter;
      keepWithNext = measured.keepWithNext ?? false;
    }

    return {
      candidates,
      opened,
      minFirstPiece,
      breakAfter,
      breakBefore,
      appliedHeight,
      keepWithNext,
    };
  },

  // Only the kind of the block a cut stands in is asked (`page/pageDecorations`), so the control
  // asked here is the top-level one and the walk goes down from it rather than up from the cut
  holdsCut(doc, at, kinds) {
    const $at = doc.resolve(at);
    if ($at.depth === 0) return false;
    const container = $at.node(1);
    return (
      isBlockControl(container) &&
      heldCut(container, $at.before(1), doc, at, kinds)
    );
  },

  decorate(pos, node, cuts, into, kinds) {
    node.forEach((child, offset, index) => {
      const childPos = pos + 1 + offset;
      const inside: PageCut[] = [];
      for (const cut of cuts) {
        if (cut.at === childPos && index > 0) {
          into.push(
            Decoration.widget(childPos, () => containerSpace(cut.height), {
              key: `container-page-space-${childPos}-${cut.height}`,
              side: -100,
            })
          );
        } else if (cut.at > childPos && cut.at < childPos + child.nodeSize) {
          inside.push(cut);
        }
      }
      blockKindFor(kinds, child).decorate(childPos, child, inside, into, kinds);
    });
  },
};
