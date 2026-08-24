/**
 * Maintains one tab stop in a composite widget. It writes `tabIndex` onto discovered controls and
 * focuses without scrolling because scrolling dismisses floating panels.
 */

import {
  type RefObject,
  useCallback,
  useLayoutEffect,
  useMemo,
  useRef,
} from "react";

export interface RovingFocusOptions {
  container: RefObject<HTMLElement | null>;
  selector: string;
  /** Limits which discovered controls can hold the tab stop. */
  navigable?: (control: HTMLElement) => boolean;
  /** Initial control; defaults to the first navigable control. */
  start?: (controls: readonly HTMLElement[]) => HTMLElement | undefined;
  /** Moves focus to the initial control when the widget appears. */
  takeFocus?: boolean;
}

export interface RovingFocus {
  controls: () => HTMLElement[];
  moveTo: (control: HTMLElement | undefined) => void;
}

export function useRovingFocus({
  container,
  selector,
  navigable,
  start,
  takeFocus = false,
}: RovingFocusOptions): RovingFocus {
  const stop = useRef<HTMLElement | null>(null);
  const took = useRef(false);

  const all = useCallback(
    () =>
      Array.from(
        container.current?.querySelectorAll<HTMLElement>(selector) ?? []
      ),
    [container, selector]
  );

  const controls = useCallback(
    () => (navigable ? all().filter(navigable) : all()),
    [all, navigable]
  );

  const write = useCallback(
    (onto: HTMLElement | null) => {
      for (const control of all()) control.tabIndex = control === onto ? 0 : -1;
      stop.current = onto;
    },
    [all]
  );

  const moveTo = useCallback(
    (control: HTMLElement | undefined) => {
      if (!control || !controls().includes(control)) return;
      write(control);
      control.focus({ preventScroll: true });
    },
    [controls, write]
  );

  // Editor-state renders can replace controls, so reconcile the tab stop after every render.
  useLayoutEffect(() => {
    const found = controls();
    if (found.length === 0) {
      write(null);
      took.current = false;
      return;
    }
    const focused = found.find((control) => control === document.activeElement);
    const held = stop.current;
    const onto =
      focused ??
      (held && found.includes(held) ? held : (start?.(found) ?? found[0])) ??
      null;
    write(onto);
    if (!takeFocus || took.current || !onto) return;
    onto.focus({ preventScroll: true });
    // Invisible panels retry on the render after placement.
    took.current = document.activeElement === onto;
  });

  return useMemo(() => ({ controls, moveTo }), [controls, moveTo]);
}
