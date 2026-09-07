/**
 * The commands and the query one run property is edited and read through, built from the
 * property's row in the run property table.
 *
 * Each edit preserves the run's untouched XML: the row says which children of the rPr change,
 * and the mark keeps every other attr it wore.
 */

import type { Mark } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import {
  type EditableRunKey,
  inheritedRunFormat,
  RUN_PROPERTIES,
  type RunSetting,
  resolveParagraph,
  resolveRun,
} from "../../../docx/formatting";
import {
  editRunProps,
  matchesRunEdit,
  type RunEdit,
} from "../../../docx/runProps";
import { docxSchema } from "../../../schema";
import { editShut, openStretches } from "../../../schema/guards";
import { documentFormatting } from "../../documentStyles";
import {
  activePieces,
  caretPiece,
  type TextPiece,
  text,
  textPieces,
} from "./shared";

function editedMark<K extends EditableRunKey>(
  state: EditorState,
  target: TextPiece,
  edit: RunEdit<K>
): Mark | null {
  const context = documentFormatting(state);
  const rPr = text(target.mark?.attrs.rPr);
  const paragraph = resolveParagraph(target.pPr, context);
  const next = editRunProps(
    { rPr, format: target.format },
    inheritedRunFormat(rPr, paragraph, context),
    edit
  );
  if (!next) return null;
  // Original values we do not read, such as rAttrs, are inherited as they are
  return docxSchema.marks.run.create({
    ...target.mark?.attrs,
    rPr: next.rPr,
    format: resolveRun(next.rPr, paragraph, context),
  });
}

interface MarkChange {
  from: number;
  to: number;
  mark: Mark;
}

/** Builds the new mark for every piece up front. If even one cannot be edited, the whole thing is abandoned */
function planChanges<K extends EditableRunKey>(
  state: EditorState,
  pieces: TextPiece[],
  edit: RunEdit<K>
): MarkChange[] | null {
  const changes: MarkChange[] = [];
  for (const target of pieces) {
    const mark = editedMark(state, target, edit);
    if (!mark) return null;
    changes.push({ from: target.from, to: target.to, mark });
  }
  return changes;
}

function applyToSelection<K extends EditableRunKey>(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  edit: RunEdit<K>
): boolean {
  const pieces = openStretches(
    state,
    // Text already in the desired state is left untouched, so its original XML survives
    textPieces(state).filter((target) => !matchesRunEdit(target.format, edit)),
    "mark"
  );
  const changes = pieces.length > 0 ? planChanges(state, pieces, edit) : null;
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

/**
 * With a collapsed caret, the formatting is only staged for the text typed next, so what settles it
 * is whether that text could go in at all.
 */
function applyToCaret<K extends EditableRunKey>(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  edit: RunEdit<K>
): boolean {
  if (editShut(state, { kind: "insert", at: state.selection.from })) {
    return false;
  }
  const target = caretPiece(state);
  if (matchesRunEdit(target.format, edit)) return false;
  const mark = editedMark(state, target, edit);
  if (!mark) return false;
  if (dispatch) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    dispatch(state.tr.setStoredMarks(mark.addToSet(marks)));
  }
  return true;
}

function runEditCommand<K extends EditableRunKey>(edit: RunEdit<K>): Command {
  return (state, dispatch) =>
    state.selection.empty
      ? applyToCaret(state, dispatch, edit)
      : applyToSelection(state, dispatch, edit);
}

export interface RunPropertyCommands<K extends EditableRunKey> {
  /** Writes the value into the selection. Null withdraws the run's own setting, so the layers below decide again */
  set(value: RunSetting<K> | null): Command;
  /** Whether the property is on where the selection stands. It counts as on only when it is on everywhere */
  isActive(state: EditorState): boolean;
  /**
   * The same toggle convention as Word: if every character in the selection is already on, turn
   * them all off; otherwise switch them all on as `on`.
   */
  toggle(on: RunSetting<K>): Command;
}

export function runPropertyCommands<K extends EditableRunKey>(
  key: K
): RunPropertyCommands<K> {
  const property = RUN_PROPERTIES[key];
  const isOn = (target: TextPiece): boolean =>
    property.isOn(target.format ?? {});
  const set = (value: RunSetting<K> | null): Command =>
    runEditCommand({ key, value });
  return {
    set,
    isActive: (state) => {
      const pieces = activePieces(state);
      return pieces.length > 0 && pieces.every(isOn);
    },
    toggle: (on) => (state, dispatch) => {
      const pieces = openStretches(
        state,
        state.selection.empty ? [caretPiece(state)] : textPieces(state),
        "mark"
      );
      return set(pieces.every(isOn) ? null : on)(state, dispatch);
    },
  };
}
