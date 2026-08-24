/**
 * Dismisses a panel on outside press, Escape, scroll, or resize. Scroll uses capture to catch
 * nested containers; events inside the panel are ignored so its own scrolling stays open.
 */

import { type RefObject, useEffect } from "react";

export function useDismiss(
  box: RefObject<HTMLElement | null>,
  open: boolean,
  onDismiss: () => void
): void {
  useEffect(() => {
    if (!open) return;

    const inside = (event: Event): boolean => {
      const target = event.target;
      return target instanceof Node && box.current?.contains(target) === true;
    };
    const onOutside = (event: Event) => {
      if (inside(event)) return;
      onDismiss();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onDismiss();
    };

    document.addEventListener("mousedown", onOutside, true);
    document.addEventListener("keydown", onKeyDown, true);
    document.addEventListener("scroll", onOutside, true);
    window.addEventListener("resize", onDismiss);

    return () => {
      document.removeEventListener("mousedown", onOutside, true);
      document.removeEventListener("keydown", onKeyDown, true);
      document.removeEventListener("scroll", onOutside, true);
      window.removeEventListener("resize", onDismiss);
    };
  }, [box, open, onDismiss]);
}
