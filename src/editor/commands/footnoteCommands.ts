/**
 * Puts a footnote into the document and writes what one says.
 *
 * The commands are built for a kind of note and exported for footnotes alone, since the footnotes
 * part is the one notes part the export writes back. Deleting and copying a reference need no
 * command: the edit that does either settles the note it calls (`editor/plugins/noteLifecycle`).
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
  // The note just put in is the one the reader is about to write, so whatever draws the notes is
  // told to open it (`editor/plugins/noteNavigation`)
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

/**
 * The body written as the story the note is held in, and null for a body nobody would take: one
 * that is not a document, one for a separator or a note the text does not call, and one already
 * saying what the note says.
 */
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
  // The reference is written again, attr for attr, so the edit reaches the place the note is
  // called from: a story stands on the document node, where no lock and no protection over that
  // text would otherwise see it
  return setStory(state.tr, key, body).setNodeMarkup(
    found.pos,
    null,
    found.node.attrs,
    found.node.marks
  );
}

/**
 * What one note of this kind says, written into the document, which is the one way an editing
 * view's edit reaches it (`editor/notes/noteSurface`).
 */
export function noteBodyCommand(
  kind: NoteKind,
  id: string,
  body: PMNode
): Command {
  return guardedCommand((state) => noteBodyTransaction(state, kind, id, body));
}

/**
 * Puts a footnote reference at the end of the selection, calling a new footnote that holds no text
 * yet, and leaves the selection standing.
 *
 * The footnote is written in the document's footnote text style and its number in the footnote
 * reference style where the document defines them, and in superscript where it does not.
 */
export const insertFootnote: Command = insertNoteCommand("footnote");

/** Whether a footnote can go in at the end of the selection, which is `insertFootnote` asked without a dispatch */
export function canInsertFootnote(state: EditorState): boolean {
  return insertFootnote(state);
}

/**
 * Replaces what one footnote says with a body of `docxSchema`, a paragraph style, a bold run and a
 * second paragraph included.
 *
 * It applies where an edit at the footnote's first reference would, and never to a separator, a
 * footnote the text does not refer to, or a body saying what the footnote already says.
 */
export function setFootnoteBody(id: string, body: PMNode): Command {
  return noteBodyCommand("footnote", id, body);
}

/**
 * Puts the caret just after a footnote's reference and opens the footnote for editing, which is
 * what a press on its number runs (`editor/plugins/noteNavigation`).
 *
 * It applies wherever the document refers to that footnote and answers the same under every mode:
 * opening a note is reading it. A footnote opened where the body is shut takes no typing.
 */
export function openFootnote(id: string): Command {
  return openNoteCommand("footnote", id);
}
