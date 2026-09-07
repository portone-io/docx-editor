/** Document-level readers for whether the document as it stands can be written back. */

import type { EditorState } from "prosemirror-state";
import { type ExportProblem, exportProblems } from "../../docx/invariants";
import { documentOf } from "../editorDocument";

export type { ExportProblem } from "../../docx/invariants";

/**
 * Every reason writing the document back would be refused, in the order `exportDocx` would raise
 * them, each with the code and message the refusal would carry and the position in the document
 * where there is one. Empty when the file would be written.
 *
 * A state built without an opened document has no file to write back into, so nothing about it
 * can be refused and it reports no problem.
 */
export function documentExportProblems(
  state: EditorState
): readonly ExportProblem[] {
  const { session } = documentOf(state);
  return session === null ? [] : exportProblems(state.doc, session);
}

/** Whether the document as it stands can be written back, which is what an export control is drawn from */
export function canExport(state: EditorState): boolean {
  return documentExportProblems(state).length === 0;
}
