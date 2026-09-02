/** Applies inline formatting while preserving each run's untouched XML. */

import type { Mark } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import {
  editRunProps,
  isRunToggleOn,
  matchesRunEdit,
  type RunEdit,
  type RunToggle,
} from "../../../docx/runProps";
import { docxSchema } from "../../../schema";
import {
  insertionInsideLocked,
  rangeTouchesLocked,
} from "../../../schema/locks";
import { editsShut } from "../../../schema/protectionState";
import {
  activePieces,
  caretPiece,
  type TextPiece,
  text,
  textPieces,
} from "./shared";

export type { RunToggle } from "../../../docx/runProps";

function editedMark(target: TextPiece, edit: RunEdit): Mark | null {
  const next = editRunProps(
    { rPr: text(target.mark?.attrs.rPr), format: target.format },
    target.pPr,
    edit
  );
  if (!next) return null;
  // Original values we do not read, such as rAttrs, are inherited as they are
  return docxSchema.marks.run.create({
    ...target.mark?.attrs,
    rPr: next.rPr,
    format: next.format,
  });
}

interface MarkChange {
  from: number;
  to: number;
  mark: Mark;
}

/** Builds the new mark for every piece up front. If even one cannot be edited, the whole thing is abandoned */
function planChanges(pieces: TextPiece[], edit: RunEdit): MarkChange[] | null {
  const changes: MarkChange[] = [];
  for (const target of pieces) {
    const mark = editedMark(target, edit);
    if (!mark) return null;
    changes.push({ from: target.from, to: target.to, mark });
  }
  return changes;
}

function applyToSelection(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  edit: RunEdit
): boolean {
  const pieces = textPieces(state).filter(
    (target) =>
      // Text already in the desired state is left untouched, so its original XML survives
      !matchesRunEdit(target.format, edit) &&
      // A locked stretch is left out rather than refused: the guard turns down the whole
      // transaction, so asking for it would leave the rest of the selection unformatted too
      !rangeTouchesLocked(state.doc, target.from, target.to)
  );
  const changes = pieces.length > 0 ? planChanges(pieces, edit) : null;
  if (!changes) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const change of changes) {
      tr.addMark(change.from, change.to, change.mark);
    }
    dispatch(tr);
  }
  return true;
}

/** With a collapsed caret, the formatting is only staged for the text typed next */
function applyToCaret(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  edit: RunEdit
): boolean {
  const target = caretPiece(state);
  if (matchesRunEdit(target.format, edit)) return false;
  const mark = editedMark(target, edit);
  if (!mark) return false;
  if (dispatch) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    dispatch(state.tr.setStoredMarks(mark.addToSet(marks)));
  }
  return true;
}

function runEditCommand(edit: RunEdit): Command {
  return (state, dispatch) => {
    if (editsShut(state)) return false;
    return state.selection.empty
      ? applyToCaret(state, dispatch, edit)
      : applyToSelection(state, dispatch, edit);
  };
}

/**
 * Whether character formatting reaches anything where the selection stands, which is what a
 * control offering it is drawn from.
 *
 * A selection running over a lock keeps its say, because the locked pieces are left out rather
 * than the whole edit refused; only a selection the lock leaves nothing of has nothing to format.
 * A caret formats no text of its own - it stages the formatting for the text typed next - so what
 * settles it is whether that text could go in at all.
 */
export function canFormatText(state: EditorState): boolean {
  if (editsShut(state)) return false;
  if (state.selection.empty) {
    return !insertionInsideLocked(state.doc, state.selection.from);
  }
  return textPieces(state).some(
    (target) => !rangeTouchesLocked(state.doc, target.from, target.to)
  );
}

/** Whether a toggled format is on at the current position. It counts as on only when it is on everywhere */
function isToggleActive(state: EditorState, toggle: RunToggle): boolean {
  const pieces = activePieces(state);
  return (
    pieces.length > 0 &&
    pieces.every((target) => isRunToggleOn(target.format, toggle))
  );
}

/**
 * The same toggle convention as Word.
 * If every character in the selection is already on, turn them all off; otherwise turn them all on.
 */
function toggleCommand(toggle: RunToggle): Command {
  return (state, dispatch) => {
    const edit: RunEdit = {
      kind: "toggle",
      toggle,
      on: !isToggleActive(state, toggle),
    };
    return runEditCommand(edit)(state, dispatch);
  };
}

export const toggleBold: Command = toggleCommand("bold");
export const toggleItalic: Command = toggleCommand("italic");
export const toggleUnderline: Command = toggleCommand("underline");
export const toggleStrike: Command = toggleCommand("strike");

/** Sets the font size in points. Null withdraws the setting and falls back to the document default */
export function setFontSize(pt: number | null): Command {
  return runEditCommand({ kind: "fontSize", pt });
}

/** Sets the font by name. Null withdraws the setting and falls back to the document default font */
export function setFontFamily(name: string | null): Command {
  return runEditCommand({ kind: "fontFamily", name });
}

/** Sets the text color as `#RRGGBB`. Null withdraws the color setting */
export function setTextColor(hex: string | null): Command {
  return runEditCommand({ kind: "color", hex });
}

/**
 * Sets the text background color as `#RRGGBB`. Null withdraws the background.
 * A highlight (`w:highlight`) written by an older document is removed along with it at that spot.
 */
export function setTextBackground(hex: string | null): Command {
  return runEditCommand({ kind: "background", hex });
}

export function isBoldActive(state: EditorState): boolean {
  return isToggleActive(state, "bold");
}

export function isItalicActive(state: EditorState): boolean {
  return isToggleActive(state, "italic");
}

export function isUnderlineActive(state: EditorState): boolean {
  return isToggleActive(state, "underline");
}

export function isStrikeActive(state: EditorState): boolean {
  return isToggleActive(state, "strike");
}
