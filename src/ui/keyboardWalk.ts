/** One-dimensional arrow/Home/End navigation backed by roving focus. */

import {
  type KeyboardEvent,
  type RefObject,
  useCallback,
  useMemo,
} from "react";
import { useRovingFocus } from "./rovingFocus";

/** Prevent these keys from scrolling a screen-anchored panel out of place. */
export const SCROLLING_KEYS = new Set(["PageUp", "PageDown"]);

export type WalkOrientation = "vertical" | "horizontal";

const WALK_KEYS: Readonly<
  Record<WalkOrientation, { previous: string; next: string }>
> = {
  vertical: { previous: "ArrowUp", next: "ArrowDown" },
  horizontal: { previous: "ArrowLeft", next: "ArrowRight" },
};

function walkedTo(
  controls: readonly HTMLElement[],
  from: number,
  key: string,
  orientation: WalkOrientation
): HTMLElement | undefined {
  const { previous, next } = WALK_KEYS[orientation];
  switch (key) {
    case next:
      return controls[(from + 1) % controls.length];
    case previous:
      return controls[(from - 1 + controls.length) % controls.length];
    case "Home":
      return controls[0];
    case "End":
      return controls[controls.length - 1];
    default:
      return undefined;
  }
}

export interface LinearWalkOptions {
  container: RefObject<HTMLElement | null>;
  selector: string;
  orientation: WalkOrientation;
  /** Filters which controls the walk may land on; all of them by default. */
  navigable?: (control: HTMLElement) => boolean;
  /** The control to open on; the first by default. */
  start?: (controls: readonly HTMLElement[]) => HTMLElement | undefined;
  takeFocus?: boolean;
}

export interface LinearWalk {
  controls: () => HTMLElement[];
  moveTo: (control: HTMLElement | undefined) => void;
  /** Handles one key; returns whether it was the walk's own so the caller can handle the rest. */
  walk: (event: KeyboardEvent) => boolean;
}

export function useLinearWalk({
  container,
  selector,
  orientation,
  navigable,
  start,
  takeFocus,
}: LinearWalkOptions): LinearWalk {
  const roving = useRovingFocus({
    container,
    selector,
    navigable,
    start,
    takeFocus,
  });

  const walk = useCallback(
    (event: KeyboardEvent) => {
      const controls = roving.controls();
      if (controls.length === 0) return false;
      const found = controls.findIndex(
        (control) => control === document.activeElement
      );
      const to = walkedTo(
        controls,
        found === -1 ? 0 : found,
        event.key,
        orientation
      );
      if (!to) return false;
      roving.moveTo(to);
      event.preventDefault();
      return true;
    },
    [roving, orientation]
  );

  return useMemo(
    () => ({ controls: roving.controls, moveTo: roving.moveTo, walk }),
    [roving, walk]
  );
}
