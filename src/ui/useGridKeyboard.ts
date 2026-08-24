/** Two-dimensional ARIA grid navigation with clamped edges and Tab dismissal. */

import { type KeyboardEvent, type RefObject, useCallback } from "react";
import { SCROLLING_KEYS } from "./keyboardWalk";
import { useRovingFocus } from "./rovingFocus";

const GRID_ROWS = '[role="row"]';
const GRID_CELLS = '[role="gridcell"]';

interface Spot {
  row: number;
  column: number;
}

function rowsOf(grid: HTMLElement | null): HTMLElement[][] {
  return Array.from(
    grid?.querySelectorAll<HTMLElement>(GRID_ROWS) ?? [],
    (row) => Array.from(row.querySelectorAll<HTMLElement>(GRID_CELLS))
  );
}

function spotOf(rows: readonly HTMLElement[][], cell: Element | null): Spot {
  for (const [row, cells] of rows.entries()) {
    const column = cells.findIndex((found) => found === cell);
    if (column !== -1) return { row, column };
  }
  return { row: 0, column: 0 };
}

function held(value: number, last: number): number {
  return Math.max(0, Math.min(value, last));
}

function stepped(
  rows: readonly HTMLElement[][],
  from: Spot,
  event: KeyboardEvent
): Spot | null {
  const acrossRows = (row: number): Spot => ({
    row,
    column: held(from.column, rows[row].length - 1),
  });
  const wholeGrid = event.ctrlKey || event.metaKey;
  switch (event.key) {
    case "ArrowRight":
      return {
        row: from.row,
        column: held(from.column + 1, rows[from.row].length - 1),
      };
    case "ArrowLeft":
      return {
        row: from.row,
        column: held(from.column - 1, rows[from.row].length - 1),
      };
    case "ArrowDown":
      return acrossRows(held(from.row + 1, rows.length - 1));
    case "ArrowUp":
      return acrossRows(held(from.row - 1, rows.length - 1));
    case "Home":
      return wholeGrid ? { row: 0, column: 0 } : { row: from.row, column: 0 };
    case "End": {
      const row = wholeGrid ? rows.length - 1 : from.row;
      return { row, column: rows[row].length - 1 };
    }
    default:
      return null;
  }
}

export interface GridKeyboardOptions {
  grid: RefObject<HTMLElement | null>;
  onClose: () => void;
  start?: (cells: readonly HTMLElement[]) => HTMLElement | undefined;
  takeFocus?: boolean;
}

export interface GridKeyboard {
  onKeyDown: (event: KeyboardEvent) => void;
}

export function useGridKeyboard({
  grid,
  onClose,
  start,
  takeFocus = true,
}: GridKeyboardOptions): GridKeyboard {
  const roving = useRovingFocus({
    container: grid,
    selector: GRID_CELLS,
    start,
    takeFocus,
  });

  const onKeyDown = useCallback(
    (event: KeyboardEvent) => {
      const rows = rowsOf(grid.current).filter((cells) => cells.length > 0);
      if (rows.length === 0) return;
      if (event.key === "Tab") {
        onClose();
        event.preventDefault();
        return;
      }
      if (SCROLLING_KEYS.has(event.key)) {
        event.preventDefault();
        return;
      }
      const to = stepped(rows, spotOf(rows, document.activeElement), event);
      if (!to) return;
      roving.moveTo(rows[to.row][to.column]);
      event.preventDefault();
    },
    [grid, onClose, roving]
  );

  return { onKeyDown };
}
