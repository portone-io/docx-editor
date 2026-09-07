/**
 * Measures the sheet as drawn on screen: the height of each body block, the gap above it, and
 * where the page breaks inside it stand, with nothing computed line by line.
 *
 * Every place a block may be parted at is read as a break candidate (`page/blockKinds`) by the
 * block's own kind (`page/kinds`), so this pass names no shape of block itself.
 * The engine's own marks are taken back off as they are read, so a block already pushed and a
 * break already given its space read as they would with neither applied. A measurement that read
 * them in would add to what is already there, and the layout would creep on every pass.
 * Where there is no layout (in tests) everything is 0, so the result is a single page.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { editorAttributes } from "../styles/classNames";
import { blockKindFor, type MeasuredBlock } from "./blockKinds";
import { blockKindsOf } from "./pageDecorations";

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

interface DrawnBlock {
  node: PMNode;
  pos: number;
  dom: HTMLElement;
}

function drawnBlocks(view: EditorView): DrawnBlock[] {
  const found: DrawnBlock[] = [];
  view.state.doc.forEach((node, offset) => {
    const dom = view.nodeDOM(offset);
    if (dom instanceof HTMLElement && dom.getBoundingClientRect().height > 0) {
      found.push({ node, pos: offset, dom });
    }
  });
  return found;
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

  const kinds = blockKindsOf(view.state);
  const blocks: MeasuredBlock[] = [];
  let previousBottom = contentTop;
  /** Everything the engine has opened up above the point being read */
  let applied = 0;

  for (const { node, pos, dom } of drawnBlocks(view)) {
    const rect = dom.getBoundingClientRect();
    applied += appliedPush(dom);
    /** Everything opened up above this block, which its own measurements are read without */
    const above = applied;
    const blockY = (viewportY: number) => sheetY(viewportY) - above;
    const top = blockY(rect.top);
    const measured = blockKindFor(kinds, node).measure({
      view,
      node,
      pos,
      dom,
      sheetY: blockY,
      top,
      scale,
    });

    applied += measured.appliedHeight;
    const bottom = sheetY(rect.bottom) - applied;
    blocks.push({
      pos,
      gap: top - previousBottom,
      height: bottom - top,
      breakBefore: dom.hasAttribute(editorAttributes.pageBreakBefore),
      breakAfter: measured.breakAfter,
      candidates: measured.candidates,
      minFirstPiece: measured.minFirstPiece,
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
