/**
 * The scale an element is drawn at, read back off the transforms standing over it.
 *
 * The page layer is scaled with a transform rather than the `zoom` property (`./editor.css`), so
 * the scale has to be read out of whatever the transform resolves to: a browser hands back the used
 * matrix, while a document implementation that does not compute one hands the function back as it
 * was written. Everything that turns a rectangle read off the screen into a measurement on the
 * paper divides by this number, so a transform an application puts above the editor - a dialog
 * opening on a scale, a container fitting itself to a small screen - counts as much as the layer's
 * own: a rectangle read inside one of those is drawn at their product.
 *
 * The drawn width over `offsetWidth` would answer without parsing anything, but `offsetWidth` is
 * rounded to whole pixels, and a scale a ten-thousandth out is a measurement half a pixel out by
 * the foot of a long document, which is the drift this scaling exists to avoid.
 */

import { editorClassNames } from "./classNames";

const HORIZONTAL_SCALE =
  /^(?:matrix3d|matrix|scale3d|scaleX|scale)\(\s*(-?\d*\.?\d+(?:e[+-]?\d+)?)/i;

function styleOf(element: Element): CSSStyleDeclaration | null {
  const view = element.ownerDocument.defaultView;
  return view ? view.getComputedStyle(element) : null;
}

/** The horizontal scale of one element's own transform. 1 for an element carrying none */
function ownScale(element: Element): number {
  const matched = HORIZONTAL_SCALE.exec(styleOf(element)?.transform ?? "");
  const scale = Number.parseFloat(matched?.[1] ?? "");
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** The horizontal scale of an element's own transform and of every transform above it */
export function visualScaleOf(element: Element | null | undefined): number {
  let scale = 1;
  for (let at = element ?? null; at !== null; at = at.parentElement) {
    scale *= ownScale(at);
  }
  return scale;
}

/** The scale the paper a node is drawn on is being read at */
export function pageScaleAround(node: Element): number {
  return visualScaleOf(node.closest(`.${editorClassNames.pageLayer}`));
}
