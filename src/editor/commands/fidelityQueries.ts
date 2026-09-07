/** Document-level reader for what the open document holds that the editor cannot model. */

import type { EditorState } from "prosemirror-state";
import { type FidelityNote, fidelityNotesOf } from "../../docx/fidelity";

export type {
  FidelityCode,
  FidelityNote,
  FidelitySeverity,
} from "../../docx/fidelity";

/**
 * What the document standing in the editor hides, stands a placeholder in front of, or approximates.
 *
 * The notes are read off the document as it is now, so a note goes away as soon as an edit removes
 * what it was about. `part` is null: an editor state carries the document and not the file it came
 * from, and `importDocx` is where a note can name the part it belongs to.
 */
export function documentFidelity(state: EditorState): readonly FidelityNote[] {
  return fidelityNotesOf(state.doc, null);
}
