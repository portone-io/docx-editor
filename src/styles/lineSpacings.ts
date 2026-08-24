/**
 * The line spacings the built-in menu offers: the four presets Word and Google Docs offer.
 * Each carries the value `setLineSpacing` takes, so a menu of your own runs the same command
 * off the same list.
 */

import type { LineSpacing } from "../model/format";

export interface LineSpacingOption {
  label: string;
  spacing: LineSpacing;
}

export const DEFAULT_LINE_SPACINGS: readonly LineSpacingOption[] = [
  { label: "Single", spacing: { rule: "auto", lines: 1 } },
  { label: "1.15", spacing: { rule: "auto", lines: 1.15 } },
  { label: "1.5", spacing: { rule: "auto", lines: 1.5 } },
  { label: "Double", spacing: { rule: "auto", lines: 2 } },
];
