import { type RefObject, useLayoutEffect, useState } from "react";
import { fitWidthZoom } from "./zoom";

export function useFitWidthZoom(
  container: RefObject<HTMLElement | null>,
  pageWidth: number
): number {
  const [zoom, setZoom] = useState(1);

  useLayoutEffect(() => {
    const element = container.current;
    if (!element) return;
    const frameWindow = element.ownerDocument.defaultView;
    const update = () => {
      setZoom(fitWidthZoom(element.clientWidth, pageWidth));
    };

    update();
    const observer =
      typeof ResizeObserver === "undefined" ? null : new ResizeObserver(update);
    observer?.observe(element);
    frameWindow?.addEventListener("resize", update, { passive: true });

    return () => {
      observer?.disconnect();
      frameWindow?.removeEventListener("resize", update);
    };
  }, [container, pageWidth]);

  return zoom;
}
