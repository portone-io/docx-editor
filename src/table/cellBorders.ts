/**
 * The cell border presets the built-in menu offers, each with the name written on its row.
 * A menu of your own takes the list from the root entry and hands the preset it picked to
 * `setCellBorders`.
 */

import type { CellBorderPreset } from "./cellFormatting";

export interface CellBorderOption {
  /** What the row is called on screen */
  label: string;
  preset: CellBorderPreset;
}

export const DEFAULT_CELL_BORDERS: readonly CellBorderOption[] = [
  { label: "All borders", preset: "all" },
  { label: "Outer borders", preset: "outer" },
  { label: "No borders", preset: "none" },
];
