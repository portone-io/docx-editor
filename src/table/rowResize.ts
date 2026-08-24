import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { editRowHeight } from "../docx/tableFormatting";
import { toRowFormat } from "../model/format";
import { isLockedCell } from "../schema/locks";

export const MIN_ROW_HEIGHT_PT = 6;

export interface RowResize {
  /** Position immediately before the table. */
  tablePos: number;
  row: number;
  heightPt: number;
}

/** Converts a pointer delta to a valid row-height floor, accounting for document zoom. */
export function resizedRowHeight(
  startPt: number,
  deltaPx: number,
  scale: number
): number {
  const effectiveScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const next = startPt + (deltaPx / effectiveScale) * (72 / 96);
  return Math.max(MIN_ROW_HEIGHT_PT, Math.round(next * 20) / 20);
}

/** Absolute document position immediately before one row in a table. */
export function rowPositionAt(
  tablePos: number,
  table: PMNode,
  row: number
): number {
  let pos = tablePos + 1;
  for (let index = 0; index < row; index += 1)
    pos += table.child(index).nodeSize;
  return pos;
}

export function buildResizeRowTransaction(
  state: EditorState,
  resize: RowResize
): Transaction | null {
  const table = state.doc.nodeAt(resize.tablePos);
  if (
    !table ||
    table.type.spec.tableRole !== "table" ||
    resize.row < 0 ||
    resize.row >= table.childCount
  ) {
    return null;
  }
  const row = table.child(resize.row);
  if (row.children.some(isLockedCell)) return null;
  const edited = editRowHeight(
    typeof row.attrs.trPr === "string" ? row.attrs.trPr : null,
    resize.heightPt
  );
  if (!edited) return null;
  return state.tr.setNodeMarkup(
    rowPositionAt(resize.tablePos, table, resize.row),
    null,
    {
      ...row.attrs,
      trPr: edited.trPr,
      format: toRowFormat(edited.format),
    }
  );
}
