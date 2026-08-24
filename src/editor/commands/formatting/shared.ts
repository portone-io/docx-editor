/** Shared text-piece discovery for formatting reads and edits. */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { type RunFormat, toRunFormat } from "../../../model/format";
import { docxSchema } from "../../../schema";
import { rangeTouchesLocked } from "../../../schema/locks";

/** One piece of text whose formatting is to be edited */
export interface TextPiece {
  from: number;
  to: number;
  /** The run mark this text carries. Null for text with no formatting */
  mark: Mark | null;
  /** The formatting XML of the paragraph the text sits in. Used to check for style inheritance */
  pPr: string | null;
  /** The effective formatting visible on screen */
  format: RunFormat | null;
}

export function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

export function runMarkOf(marks: readonly Mark[]): Mark | null {
  return marks.find((mark) => mark.type === docxSchema.marks.run) ?? null;
}

function piece(
  from: number,
  to: number,
  marks: readonly Mark[],
  parent: PMNode | null
): TextPiece {
  const mark = runMarkOf(marks);
  return {
    from,
    to,
    mark,
    pPr: text(parent?.attrs.pPr),
    format: toRunFormat(mark?.attrs.format),
  };
}

/**
 * The text pieces inside the selection.
 * Non-text content is skipped. A tab is text in the model and shares its OOXML run properties,
 * so character formatting reaches it along with the visible characters around it.
 * Selections with several ranges, such as a table cell selection, go down the same path.
 */
export function textPieces(state: EditorState): TextPiece[] {
  const pieces: TextPiece[] = [];
  for (const range of state.selection.ranges) {
    const from = range.$from.pos;
    const to = range.$to.pos;
    state.doc.nodesBetween(from, to, (node, pos, parent) => {
      if (!node.isText) return true;
      const start = Math.max(pos, from);
      const end = Math.min(pos + node.nodeSize, to);
      if (start < end) pieces.push(piece(start, end, node.marks, parent));
      return false;
    });
  }
  return pieces;
}

/** When the selection is collapsed, the formatting that text typed next will follow */
export function caretPiece(state: EditorState): TextPiece {
  const $from = state.selection.$from;
  const marks = state.storedMarks ?? $from.marks();
  return piece($from.pos, $from.pos, marks, $from.parent);
}

/**
 * The places formatting currently applies to: a single caret, or the text pieces of the selection
 * a lock leaves open. It has to leave out what `applyToSelection` leaves out, or a toggle reads off
 * for text that is already on and can never turn it off.
 */
export function activePieces(state: EditorState): TextPiece[] {
  if (state.selection.empty) return [caretPiece(state)];
  return textPieces(state).filter(
    (target) => !rangeTouchesLocked(state.doc, target.from, target.to)
  );
}
