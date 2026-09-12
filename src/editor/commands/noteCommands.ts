/**
 * Puts a note into the document and writes what one says, exported once per kind the way
 * `toggleBold` and `toggleItalic` are rather than as one command taking a kind. Deleting and
 * copying a reference need no command: the edit that does either settles the note it calls
 * (`editor/plugins/noteLifecycle`).
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import {
  newNoteStory,
  nextNoteId,
  noteNumberRunProps,
  takenNoteIds,
} from "../../docx/notes/newNote";
import { sameStory, setStory, storyKey, storyOf } from "../../docx/story";
import { docxSchema } from "../../schema";
import { guardedCommand } from "../../schema/guards";
import type { NoteKind } from "../../schema/stories";
import { isWrapperType } from "../../schema/wrappers";
import { documentOf } from "../editorDocument";
import {
  noteReferenceAt,
  openNoteCommand,
  requestNote,
} from "../plugins/noteNavigation";

/** The wrappers the content just before this position stands inside, which a reference put there stands inside too */
function wrappersBefore(doc: PMNode, pos: number): readonly Mark[] {
  const $pos = doc.resolve(pos);
  return ($pos.nodeBefore?.marks ?? $pos.marks()).filter((mark) =>
    isWrapperType(mark.type)
  );
}

function insertNoteTransaction(
  state: EditorState,
  kind: NoteKind
): Transaction | null {
  const at = state.selection.to;
  const $at = state.doc.resolve(at);
  const type = docxSchema.nodes.noteReference;
  if (!$at.parent.canReplaceWith($at.index(), $at.index(), type)) return null;
  const { formatting, reservedNoteKeys } = documentOf(state);
  const id = nextNoteId(takenNoteIds(state.doc, kind, reservedNoteKeys));
  const reference = type.create({ kind, id }, null, [
    docxSchema.marks.run.create({ rPr: noteNumberRunProps(kind, formatting) }),
    ...wrappersBefore(state.doc, at),
  ]);
  const key = storyKey(kind, id);
  return requestNote(
    setStory(
      state.tr.insert(at, reference),
      key,
      newNoteStory(kind, formatting)
    ),
    key
  );
}

function insertNoteCommand(kind: NoteKind): Command {
  return guardedCommand((state) => insertNoteTransaction(state, kind));
}

function noteBodyTransaction(
  state: EditorState,
  kind: NoteKind,
  id: string,
  body: PMNode
): Transaction | null {
  const key = storyKey(kind, id);
  if (body.type !== docxSchema.nodes.doc) return null;
  if (documentOf(state).specialNotes.has(key)) return null;
  const found = noteReferenceAt(state.doc, kind, id);
  if (found === null) return null;
  if (sameStory(storyOf(state.doc, key), body)) return null;
  // The reference is rewritten attr for attr so the guards judge the edit at the place the note is
  // called from: a story stands on the document node, out of reach of a lock or a protection over
  // that text
  return setStory(state.tr, key, body).setNodeMarkup(
    found.pos,
    null,
    found.node.attrs,
    found.node.marks
  );
}

/** Writes what one note of this kind says; an editing view's edit reaches the document this way */
export function noteBodyCommand(
  kind: NoteKind,
  id: string,
  body: PMNode
): Command {
  return guardedCommand((state) => noteBodyTransaction(state, kind, id, body));
}

/**
 * Puts a footnote reference at the end of the selection, calling a new, empty footnote, and asks
 * for that footnote to be opened. The footnote is written in the document's footnote styles where
 * it defines them, and in superscript where it does not.
 */
export const insertFootnote: Command = insertNoteCommand("footnote");

/** Whether a footnote can go in at the end of the selection */
export function canInsertFootnote(state: EditorState): boolean {
  return insertFootnote(state);
}

/**
 * Replaces what one footnote says with a `doc` node of `docxSchema`. It applies where an edit at
 * the footnote's first reference would, and never to a separator, a footnote the text does not
 * refer to, or a body saying what the footnote already says.
 */
export function setFootnoteBody(id: string, body: PMNode): Command {
  return noteBodyCommand("footnote", id, body);
}

/**
 * Puts the caret just after a footnote's reference and opens the footnote for editing, which is
 * what a press on its number runs. It applies under every mode, since opening a note is reading
 * it; a footnote opened where the body is shut takes no typing.
 */
export function openFootnote(id: string): Command {
  return openNoteCommand("footnote", id);
}

/**
 * Puts an endnote reference at the end of the selection, calling a new, empty endnote, and asks
 * for that endnote to be opened; `insertFootnote` for an endnote, which is drawn after the last
 * paragraph of the document rather than at the foot of a page.
 */
export const insertEndnote: Command = insertNoteCommand("endnote");

/** Whether an endnote can go in at the end of the selection */
export function canInsertEndnote(state: EditorState): boolean {
  return insertEndnote(state);
}

/** Replaces what one endnote says with a `doc` node of `docxSchema`, as `setFootnoteBody` does */
export function setEndnoteBody(id: string, body: PMNode): Command {
  return noteBodyCommand("endnote", id, body);
}

/**
 * Puts the caret just after an endnote's reference and opens the endnote for editing;
 * `openFootnote` for an endnote, which stands at the end of the document, so this is also what
 * takes a reader to it.
 */
export function openEndnote(id: string): Command {
  return openNoteCommand("endnote", id);
}
