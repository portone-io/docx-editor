/**
 * Measures the sheet as drawn on screen: the height of each body block, the gap above it, and
 * where the page breaks inside it stand, with nothing computed line by line.
 *
 * Every place a block may be parted at is read as a break candidate (`page/blockKinds`), whatever
 * shape of block it is: a page break for a paragraph, a safe row boundary for a table.
 * The engine's own marks are taken back off as they are read, so a block already pushed and a
 * break already given its space read as they would with neither applied. A measurement that read
 * them in would add to what is already there, and the layout would creep on every pass.
 * Where there is no layout (in tests) everything is 0, so the result is a single page.
 */

import type { EditorView } from "prosemirror-view";
import { editorAttributes } from "../styles/classNames";
import type { BreakCandidate, MeasuredBlock } from "./blockKinds";
import { pageBreaksIn } from "./kinds/paragraphKind";
import { tableKind } from "./kinds/tableKind";

/**
 * The measurements taken in order to draw the page overlay. Positions are relative to
 * the overlay box
 */
export interface SheetMeasure {
  left: number;
  top: number;
  width: number;
  /** The sheet's top and bottom padding */
  contentTop: number;
  contentBottom: number;
  blocks: MeasuredBlock[];
}

const PAGE_BREAK_BR = `br[${editorAttributes.breakType}="page"]`;
const BREAK_SPACE = `[${editorAttributes.pageBreakSpace}]`;

function pixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function visualScale(element: HTMLElement): number {
  const scale = pixels(getComputedStyle(element).zoom);
  return scale > 0 ? scale : 1;
}

function appliedPush(element: HTMLElement): number {
  return pixels(element.getAttribute(editorAttributes.pagePush) ?? "");
}

function drawnBlocks(view: EditorView): { pos: number; dom: HTMLElement }[] {
  const found: { pos: number; dom: HTMLElement }[] = [];
  view.state.doc.forEach((_node, offset) => {
    const dom = view.nodeDOM(offset);
    if (dom instanceof HTMLElement && dom.getBoundingClientRect().height > 0) {
      found.push({ pos: offset, dom });
    }
  });
  return found;
}

/**
 * A break with no space of its own sits somewhere the space could not be opened, inside a table
 * cell (`page/pageDecorations`). Those still start a new page, but only after the whole block,
 * which the layout answers from `breakAfter`.
 */
function breakWithoutSpace(dom: HTMLElement, spaces: number): boolean {
  return dom.querySelectorAll(PAGE_BREAK_BR).length > spaces;
}

export function measureSheet(
  view: EditorView,
  layer: HTMLElement
): SheetMeasure {
  const sheet = view.dom;
  const sheetRect = sheet.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  const style = getComputedStyle(sheet);
  const scale = visualScale(layer);
  const contentTop = pixels(style.paddingTop);
  const contentBottom = pixels(style.paddingBottom);
  const sheetY = (viewportY: number) => (viewportY - sheetRect.top) / scale;

  const blocks: MeasuredBlock[] = [];
  let previousBottom = contentTop;
  /** Everything the engine has opened up above the point being read */
  let applied = 0;

  for (const { pos, dom } of drawnBlocks(view)) {
    const node = view.state.doc.nodeAt(pos);
    const rect = dom.getBoundingClientRect();
    applied += appliedPush(dom);
    /** Everything opened up above this block, which its own measurements are read without */
    const above = applied;
    const blockY = (viewportY: number) => sheetY(viewportY) - above;
    const top = blockY(rect.top);
    const measuredTable =
      node && tableKind.matches(node)
        ? tableKind.measure({
            view,
            node,
            pos,
            dom,
            sheetY: blockY,
            top,
            scale,
          })
        : null;

    // Each space element is the one the break at the same ordinal was given
    const breaks = node ? pageBreaksIn(node, pos) : [];
    const spaces = Array.from(dom.querySelectorAll(BREAK_SPACE));
    const forced: BreakCandidate[] = [];
    spaces.forEach((space, index) => {
      const box = space.getBoundingClientRect();
      const found = breaks.at(index);
      if (found) {
        forced.push({
          at: found.at,
          offset: sheetY(box.top) - applied - top,
          forced: true,
          repeatHeight: 0,
        });
      }
      applied += box.height / scale;
    });

    applied += measuredTable?.appliedHeight ?? 0;
    const bottom = sheetY(rect.bottom) - applied;
    const height = bottom - top;
    blocks.push({
      pos,
      gap: top - previousBottom,
      height,
      breakBefore: dom.hasAttribute(editorAttributes.pageBreakBefore),
      breakAfter:
        measuredTable?.breakAfter ?? breakWithoutSpace(dom, spaces.length),
      // A space is never opened inside a table, so at most one of the two lists holds anything
      candidates: [...forced, ...(measuredTable?.candidates ?? [])],
      minFirstPiece:
        measuredTable?.minFirstPiece ?? forced[0]?.offset ?? height,
    });
    previousBottom = bottom;
  }

  return {
    left: (sheetRect.left - layerRect.left) / scale,
    top: (sheetRect.top - layerRect.top) / scale,
    width: sheetRect.width / scale,
    contentTop,
    contentBottom,
    blocks,
  };
}
