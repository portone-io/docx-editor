/**
 * The public surface of table editing.
 *
 * Only what the toolbar needs is exported. Following the ProseMirror convention,
 * calling a command without `dispatch` only reports whether it can run right now, and that answer
 * takes in the lock: a structural edit that would carry a locked cell off reports that it does not
 * apply, rather than being refused after the click (`schema/locks`).
 *
 * Everything here works on a table already in the document; putting a new one into the body is
 * `insertTable` in `./commands`, which reads the page geometry the editor holds.
 */

export type {
  ActiveCellPadding,
  CellBorderPreset,
  MixedCellValue,
} from "./cellFormatting";
export {
  activeCellBackground,
  activeCellBorderColor,
  activeCellPadding,
  activeCellVerticalAlign,
  canSetCellBorderColor,
  canSetCellFormatting,
  setCellBackground,
  setCellBorderColor,
  setCellBorders,
  setCellPadding,
  setCellVerticalAlign,
} from "./cellFormatting";
export type { TableCommand } from "./commands";
export {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  splitCell,
} from "./commands";
export { canMergeCells, canSplitCell } from "./merge";
