/** ARIA menu navigation with roving focus, typeahead, and Tab dismissal. */

import { type KeyboardEvent, type RefObject, useCallback } from "react";
import { SCROLLING_KEYS, useLinearWalk } from "./keyboardWalk";

const MENU_ROWS =
  '[role="menuitem"],[role="menuitemradio"],[role="menuitemcheckbox"]';

function currentChoice(rows: readonly HTMLElement[]): HTMLElement | undefined {
  return rows.find((row) => row.getAttribute("aria-checked") === "true");
}

function typedAhead(
  rows: readonly HTMLElement[],
  from: number,
  letter: string
): HTMLElement | undefined {
  const wanted = letter.toLowerCase();
  for (let step = 1; step <= rows.length; step += 1) {
    const row = rows[(from + step) % rows.length];
    if (row.textContent?.trim().toLowerCase().startsWith(wanted)) return row;
  }
  return undefined;
}

export interface MenuKeyboardOptions {
  menu: RefObject<HTMLElement | null>;
  onClose: () => void;
  /** Whether focus moves into the menu when it appears. */
  takeFocus?: boolean;
}

export interface MenuKeyboard {
  onKeyDown: (event: KeyboardEvent) => void;
}

export function useMenuKeyboard({
  menu,
  onClose,
  takeFocus = true,
}: MenuKeyboardOptions): MenuKeyboard {
  const rows = useLinearWalk({
    container: menu,
    selector: MENU_ROWS,
    orientation: "vertical",
    start: currentChoice,
    takeFocus,
  });

  const onKeyDown = useCallback(
    (event: KeyboardEvent) => {
      const found = rows.controls();
      if (found.length === 0) return;
      if (rows.walk(event)) return;

      if (event.key === "Tab") {
        onClose();
        event.preventDefault();
        return;
      }
      if (SCROLLING_KEYS.has(event.key)) {
        event.preventDefault();
        return;
      }
      // Enter and Space are the row's own: a button answers both with a click
      const typed = event.key.length === 1 && event.key !== " ";
      if (!typed || event.altKey || event.ctrlKey || event.metaKey) return;
      const at = found.findIndex((row) => row === document.activeElement);
      const row = typedAhead(found, at === -1 ? 0 : at, event.key);
      if (!row) return;
      rows.moveTo(row);
      event.preventDefault();
    },
    [rows, onClose]
  );

  return { onKeyDown };
}
