/**
 * Positions fixed panels within the viewport so host overflow cannot clip them. Panels that do
 * not remeasure themselves are dismissed on scroll by `useDismiss`.
 */

import type { EditorView } from "prosemirror-view";
import { type RefObject, useLayoutEffect, useState } from "react";

const EDGE_GAP = 8;

const CONTROL_GAP = 4;

export interface PanelPosition {
  left: number;
  top: number;
}

interface Size {
  width: number;
  height: number;
}

interface Viewport {
  width: number;
  height: number;
}

export interface ControlBox {
  left: number;
  top: number;
  bottom: number;
}

export interface ScreenPoint {
  clientX: number;
  clientY: number;
}

/**
 * Slides a run of `length` back inside `0..limit`, keeping the edge gap at both ends.
 * A panel longer than the screen is pinned to the near edge, so what stays visible is
 * its head rather than its tail.
 */
function slideInside(start: number, length: number, limit: number): number {
  return Math.max(EDGE_GAP, Math.min(start, limit - EDGE_GAP - length));
}

/**
 * The panel hangs under the control, lined up with its left edge.
 * When it does not fit under the control but does fit over it, it flips above instead;
 * when it fits neither way it slides in from the bottom edge of the screen.
 */
export function positionUnderControl(
  control: ControlBox,
  size: Size,
  viewport: Viewport
): PanelPosition {
  const under = control.bottom + CONTROL_GAP;
  const over = control.top - CONTROL_GAP - size.height;
  const fitsUnder = under + size.height + EDGE_GAP <= viewport.height;
  return {
    left: slideInside(control.left, size.width, viewport.width),
    top:
      fitsUnder || over < EDGE_GAP
        ? slideInside(under, size.height, viewport.height)
        : over,
  };
}

export function positionAtPoint(
  point: ScreenPoint,
  size: Size,
  viewport: Viewport
): PanelPosition {
  return {
    left: slideInside(point.clientX, size.width, viewport.width),
    top: slideInside(point.clientY, size.height, viewport.height),
  };
}

/**
 * The bottom left corner of what is drawn at this position in the document, which is where a panel
 * about that spot hangs from.
 *
 * A position the browser has not laid out has no coordinates to ask for, and the panel then hangs
 * against the top of the paper rather than not at all.
 */
export function pointBelowPos(view: EditorView, pos: number): ScreenPoint {
  try {
    const at = view.coordsAtPos(pos);
    return { clientX: at.left, clientY: at.bottom };
  } catch {
    const paper = view.dom.getBoundingClientRect();
    return { clientX: paper.left, clientY: paper.top };
  }
}

function measure(panel: HTMLElement): { size: Size; viewport: Viewport } {
  return {
    size: { width: panel.offsetWidth, height: panel.offsetHeight },
    viewport: { width: window.innerWidth, height: window.innerHeight },
  };
}

/**
 * Places a panel under the control it opens from.
 *
 * The panel has to be measured, so this is decided right after it renders, before the
 * browser paints. It is `null` for that one render, which the panel spends invisible.
 */
export function usePanelUnderControl(
  panel: RefObject<HTMLElement | null>,
  control: RefObject<HTMLElement | null>,
  open: boolean
): PanelPosition | null {
  const [position, setPosition] = useState<PanelPosition | null>(null);

  useLayoutEffect(() => {
    const box = panel.current;
    const from = control.current;
    if (!open || !box || !from) {
      setPosition(null);
      return;
    }
    const { size, viewport } = measure(box);
    setPosition(
      positionUnderControl(from.getBoundingClientRect(), size, viewport)
    );
  }, [panel, control, open]);

  return position;
}

export function usePanelAtPoint(
  panel: RefObject<HTMLElement | null>,
  point: ScreenPoint
): PanelPosition | null {
  const [position, setPosition] = useState<PanelPosition | null>(null);
  const { clientX, clientY } = point;

  useLayoutEffect(() => {
    const box = panel.current;
    if (!box) return;
    const { size, viewport } = measure(box);
    setPosition(positionAtPoint({ clientX, clientY }, size, viewport));
  }, [panel, clientX, clientY]);

  return position;
}
