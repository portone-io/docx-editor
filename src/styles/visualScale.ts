/**
 * The scale the page layer is drawn at, read back off the element that carries it.
 *
 * The layer is scaled with a transform rather than the `zoom` property (`./editor.css`), so the
 * scale has to be read out of whatever the transform resolves to: a browser hands back the used
 * matrix, while a document implementation that does not compute one hands the function back as it
 * was written. Everything that turns a rectangle read off the screen into a measurement on the
 * paper divides by this number.
 */

import { editorClassNames } from "./classNames";

const HORIZONTAL_SCALE =
  /^(?:matrix3d|matrix|scale3d|scaleX|scale)\(\s*(-?\d*\.?\d+(?:e[+-]?\d+)?)/i;

function styleOf(element: Element): CSSStyleDeclaration | null {
  const view = element.ownerDocument.defaultView;
  return view ? view.getComputedStyle(element) : null;
}

/** The horizontal scale of an element's own transform. 1 for an element carrying none */
export function visualScaleOf(element: Element | null | undefined): number {
  if (!element) return 1;
  const matched = HORIZONTAL_SCALE.exec(styleOf(element)?.transform ?? "");
  const scale = Number.parseFloat(matched?.[1] ?? "");
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** The scale the paper a node is drawn on is being read at */
export function pageScaleAround(node: Element): number {
  return visualScaleOf(node.closest(`.${editorClassNames.pageLayer}`));
}
