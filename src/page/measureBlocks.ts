/**
 * Measures the sheet as drawn on screen: the height of each body block, the gap above it, and
 * where the page breaks inside it stand, with nothing computed line by line.
 *
 * The engine's own marks are taken back off as they are read, so a block already pushed and a
 * break already given its space read as they would with neither applied. A measurement that read
 * them in would add to what is already there, and the layout would creep on every pass.
 * Where there is no layout (in tests) everything is 0, so the result is a single page.
 */

import type { EditorView } from "prosemirror-view";
import { editorAttributes } from "../styles/classNames";
import type { MeasuredBlock } from "./pageLayout";
import { measureTable } from "./tableMeasurements";

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
 * cell (`page/pageDecorations`). Those still start a new page, but only after the whole block.
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
  let breakAfterPrevious = false;

  for (const { pos, dom } of drawnBlocks(view)) {
    const node = view.state.doc.nodeAt(pos);
    const rect = dom.getBoundingClientRect();
    applied += appliedPush(dom);
    const top = sheetY(rect.top) - applied;
    const measuredTable = node
      ? measureTable(view, node, pos, dom, scale)
      : null;

    const breaks: number[] = [];
    const spaces = Array.from(dom.querySelectorAll(BREAK_SPACE));
    for (const space of spaces) {
      const box = space.getBoundingClientRect();
      breaks.push(sheetY(box.top) - applied - top);
      applied += box.height / scale;
    }

    applied += measuredTable?.appliedHeight ?? 0;
    const bottom = sheetY(rect.bottom) - applied;
    blocks.push({
      pos,
      gap: top - previousBottom,
      height: bottom - top,
      breakBefore:
        breakAfterPrevious ||
        dom.hasAttribute(editorAttributes.pageBreakBefore),
      breaks,
      ...(measuredTable ? { table: measuredTable.table } : {}),
    });
    previousBottom = bottom;
    breakAfterPrevious = breakWithoutSpace(dom, spaces.length);
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
