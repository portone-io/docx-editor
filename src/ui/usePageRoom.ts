import { useLayoutEffect } from "react";

/**
 * Holds the room the page layer takes in the scroll box.
 *
 * The layer is scaled with a transform and stands outside the flow (`styles/editor.css`), so
 * nothing in the flow reaches down to where the paper ends and the scroll box would stop short of
 * it. The box the layer stands in is given the height the layer is drawn at, and the scroll box
 * then reaches the paper and keeps its own padding below it.
 *
 * The two elements are taken rather than refs to them, so that a surface which drew neither - a
 * file that was refused, before the reader opens one that is not - is measured the moment they
 * arrive rather than never again.
 */
export function usePageRoom(
  box: HTMLElement | null,
  layer: HTMLElement | null,
  scale: number
): void {
  // biome-ignore lint/correctness/useExhaustiveDependencies: the scale is not read here, since the height is taken off the layer as drawn and that rectangle already carries it. It is watched because a transform moves no box a resize observation would see, so nothing else would ask for the height again when the reader zooms
  useLayoutEffect(() => {
    if (!box || !layer) return;
    const frameWindow = box.ownerDocument.defaultView;
    let frame = 0;
    /** The height the layer is drawn at, which is the room the scroll box has to give it */
    const hold = () => {
      const next = `${layer.getBoundingClientRect().height}px`;
      if (box.style.height !== next) box.style.height = next;
    };
    // The height is written on the frame after an observation rather than inside it: a write from
    // inside dirties the layout the browser has just measured, and Chromium reports the observation
    // that raises as one it could not deliver. The height is read when that frame runs, so several
    // observations in one frame come to one reading of where the layer now stands
    const schedule = () => {
      if (frame !== 0 || !frameWindow) return;
      frame = frameWindow.requestAnimationFrame(() => {
        frame = 0;
        hold();
      });
    };

    hold();
    const observer =
      typeof ResizeObserver === "undefined"
        ? null
        : new ResizeObserver(schedule);
    observer?.observe(layer);

    return () => {
      observer?.disconnect();
      if (frame !== 0) frameWindow?.cancelAnimationFrame(frame);
      box.style.removeProperty("height");
    };
  }, [box, layer, scale]);
}
