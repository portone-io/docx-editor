import { type RefObject, useLayoutEffect } from "react";

/**
 * Holds the room the page layer takes in the scroll box.
 *
 * The layer is scaled with a transform and stands outside the flow (`styles/editor.css`), so
 * nothing in the flow reaches down to where the paper ends and the scroll box would stop short of
 * it. The box the layer stands in is given the height the layer is drawn at, and the scroll box
 * then reaches the paper and keeps its own padding below it.
 */
export function usePageRoom(
  box: RefObject<HTMLElement | null>,
  layer: RefObject<HTMLElement | null>,
  scale: number
): void {
  useLayoutEffect(() => {
    const room = box.current;
    const drawn = layer.current;
    if (!room || !drawn) return;
    const frameWindow = room.ownerDocument.defaultView;
    let frame = 0;
    /** `height` is the layer's own, which the scale has still to be applied to */
    const hold = (height: number) => {
      const next = `${height * scale}px`;
      if (room.style.height !== next) room.style.height = next;
    };

    hold(drawn.getBoundingClientRect().height / scale);
    // The layer's own box follows the text; the scale it is drawn at is watched by this effect,
    // since a transform moves no box an observer would see.
    // The height is written on the frame after the observation rather than inside it: a write from
    // inside dirties the layout the browser has just measured, and it reports the observation that
    // raises as one it could not deliver
    const observer =
      typeof ResizeObserver === "undefined" || !frameWindow
        ? null
        : new ResizeObserver(([entry]) => {
            const size = entry?.borderBoxSize?.[0];
            if (!size || frame !== 0) return;
            const height = size.blockSize;
            frame = frameWindow.requestAnimationFrame(() => {
              frame = 0;
              hold(height);
            });
          });
    observer?.observe(drawn);

    return () => {
      observer?.disconnect();
      if (frame !== 0) frameWindow?.cancelAnimationFrame(frame);
      room.style.removeProperty("height");
    };
  }, [box, layer, scale]);
}
