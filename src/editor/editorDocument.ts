/**
 * What the opened document laid down, held as one value: the formatting context every display
 * value is resolved against, the paper the body is written on, and the identifiers its comment
 * parts have already spent.
 *
 * Commands and the toolbar know nothing about the session, so the snapshot stands for it on their
 * behalf, and `editorDocumentOf` is the one place a session is read into editor values: a screen
 * and a test that open the same document hold the same values.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { DEFAULT_TAB_STOP_PT } from "../docx/documentSettings";
import {
  type FormattingContext,
  NO_DOCUMENT_DEFAULTS,
  NO_FORMATTING,
  type ParagraphStyleOption,
} from "../docx/formatting";
import { canDefineNewList } from "../docx/newLists";
import {
  DEFAULT_NOTE_NUMBERING,
  type NoteNumbering,
} from "../docx/notes/reading";
import { A4_PORTRAIT, type PageGeometry } from "../docx/pageGeometry";
import type { SessionStore } from "../docx/session";
import type { DocumentDefaults } from "../model/format";
import { NEW_LISTS_ATTR, newListsOf } from "../numbering/listRegistry";
import {
  EDITABLE_NOTE_KINDS,
  STORIES_ATTR,
  type StoryKey,
} from "../schema/stories";

/** The document-level values one editing state is built on */
export interface EditorDocument {
  /** The document these values were read from. Null for a state built without one */
  readonly session: SessionStore | null;
  /** The style chain, the defaults and the list definitions a paragraph's and a run's display values are resolved against */
  readonly formatting: FormattingContext;
  readonly defaults: DocumentDefaults;
  /** The styles the document defines for the style picker to offer */
  readonly paragraphStyles: ParagraphStyleOption[];
  /**
   * Whether the definition of a new list has somewhere to go, which is what the commands that
   * start one ask. A state built without opening a document assumes it has and behaves as every
   * state did before the question was asked.
   */
  readonly canStartNewList: boolean;
  /** The paper the document is written on, which decides the width a new table is fitted to */
  readonly geometry: PageGeometry;
  /**
   * Whether these values stand for a side story rather than for the body, whose paper is then the
   * geometry above rather than anything the story's own blocks say (`editor/documentStyles`).
   */
  readonly sideStory: boolean;
  /** The distance between automatic tab stops when no custom stop applies */
  readonly defaultTabStopPt: number;
  /** Every id already present in the opened Comments part, including unreferenced entries */
  readonly reservedCommentIds: ReadonlySet<string>;
  /** Every paragraph id present in the opened comment parts, including orphan extension entries */
  readonly reservedCommentParaIds: ReadonlySet<string>;
  /** How the document counts its footnotes and endnotes, which the note labels are spelled in */
  readonly noteNumbering: NoteNumbering;
  /** The note entries that lay out the page rather than number a note: separators and the continuation notice */
  readonly specialNotes: ReadonlySet<StoryKey>;
  /**
   * Every entry of an editable kind the opened notes parts hold, separators and entries no
   * reference names included, which a new note may not take the id of
   */
  readonly reservedNoteKeys: ReadonlySet<StoryKey>;
}

const NO_IDS: ReadonlySet<string> = new Set();

const NO_NOTE_KEYS: ReadonlySet<StoryKey> = new Set();

/** A document whose paragraph styles the editor does not know offers none to pick from */
const NO_PARAGRAPH_STYLES: ParagraphStyleOption[] = [];

/** What a state built without an opened document answers: the same as nothing being written down */
export const NO_DOCUMENT: EditorDocument = {
  session: null,
  formatting: NO_FORMATTING,
  defaults: NO_DOCUMENT_DEFAULTS,
  paragraphStyles: NO_PARAGRAPH_STYLES,
  canStartNewList: true,
  geometry: A4_PORTRAIT,
  sideStory: false,
  defaultTabStopPt: DEFAULT_TAB_STOP_PT,
  reservedCommentIds: NO_IDS,
  reservedCommentParaIds: NO_IDS,
  noteNumbering: DEFAULT_NOTE_NUMBERING,
  specialNotes: NO_NOTE_KEYS,
  reservedNoteKeys: NO_NOTE_KEYS,
};

/**
 * Every paragraph id the opened comment parts spent: the ones the comments carry and the ones
 * only the extension part names, so a new comment never takes an id an orphan entry still owns.
 */
function reservedParaIds(session: SessionStore): Set<string> {
  const paraIds = new Set<string>();
  for (const comment of session.comments.ordered) {
    if (comment.paraId !== null) paraIds.add(comment.paraId);
  }
  for (const extension of session.comments.extendedOrdered) {
    paraIds.add(extension.paraId);
  }
  return paraIds;
}

/**
 * The list definitions the editor resolves against: the ones the document wrote down, and the ones
 * the lists started while editing were registered with, which the document node carries.
 */
function formattingWithNewLists(
  formatting: FormattingContext,
  doc: PMNode
): FormattingContext {
  const added = newListsOf(doc.attrs[NEW_LISTS_ATTR]);
  if (added.size === 0 && formatting.numbering.added.size === 0) {
    return formatting;
  }
  return { ...formatting, numbering: { ...formatting.numbering, added } };
}

/** The one place an opened document is read into the values the editor holds */
export function editorDocumentOf(
  session: SessionStore,
  doc: PMNode
): EditorDocument {
  return {
    session,
    formatting: formattingWithNewLists(session.formatting, doc),
    defaults: session.defaults,
    paragraphStyles: session.paragraphStyles,
    canStartNewList: canDefineNewList(session),
    geometry: session.geometry,
    sideStory: false,
    defaultTabStopPt: session.defaultTabStopPt,
    reservedCommentIds: new Set(session.comments.byId.keys()),
    reservedCommentParaIds: reservedParaIds(session),
    noteNumbering: session.noteNumbering,
    specialNotes: session.specialNotes,
    reservedNoteKeys: new Set(
      Array.from(session.stories.values()).flatMap((story) =>
        EDITABLE_NOTE_KINDS.some((kind) => kind === story.kind)
          ? [story.key]
          : []
      )
    ),
  };
}

const editorDocumentKey = new PluginKey<EditorDocument>("docxEditorDocument");

/** The document the state was built for. `NO_DOCUMENT` for a state built without one */
export function documentOf(state: EditorState): EditorDocument {
  return editorDocumentKey.getState(state) ?? NO_DOCUMENT;
}

/**
 * The snapshot after a document-level edit.
 *
 * An opened document is read again from its session, so a snapshot handed in over the session's own
 * values stands only until the first such edit. A state built without one has no session to read,
 * and keeps the values it was given with the register alone taken from the node.
 */
function derived(current: EditorDocument, doc: PMNode): EditorDocument {
  return current.session === null
    ? {
        ...current,
        formatting: formattingWithNewLists(current.formatting, doc),
      }
    : editorDocumentOf(current.session, doc);
}

/**
 * Whether the only thing that moved is what the stories say.
 *
 * Nothing the snapshot holds is read out of a story, and typing inside one writes a story on every
 * keystroke, so a snapshot derived again there would throw away every value cached against its
 * identity - the sheet's style among them - once per key pressed inside a note. Every other
 * document-level step derives it again as it always did, one writing back the value that already
 * stood included.
 */
function onlyStoriesMoved(next: PMNode, before: PMNode): boolean {
  const names = Object.keys(next.attrs);
  return (
    next.attrs[STORIES_ATTR] !== before.attrs[STORIES_ATTR] &&
    names.length === Object.keys(before.attrs).length &&
    names.every(
      (name) => name === STORIES_ATTR || next.attrs[name] === before.attrs[name]
    )
  );
}

/**
 * Holds the snapshot for the lifetime of the state.
 *
 * It is derived again only where a document-level edit could have been recorded - the attrs of the
 * document node, which is where such an edit is written - so an ordinary edit and a selection move
 * both leave the same object behind, allowing the sheet's style to be cached against that identity.
 */
export function editorDocument(
  document: EditorDocument
): Plugin<EditorDocument> {
  return new Plugin<EditorDocument>({
    key: editorDocumentKey,
    state: {
      init: () => document,
      apply: (tr, current, old) =>
        tr.doc.attrs === old.doc.attrs || onlyStoriesMoved(tr.doc, old.doc)
          ? current
          : derived(current, tr.doc),
    },
  });
}

/**
 * The snapshot one side story is edited against: the document's own values, with the two a story
 * answers differently.
 *
 * A story has nowhere of its own to write a list definition, and the paper its text wraps at is
 * the paper of the place it is called from rather than anything the story says.
 */
export function storyDocument(
  document: EditorDocument,
  geometry: PageGeometry
): EditorDocument {
  return { ...document, sideStory: true, canStartNewList: false, geometry };
}

/** The comment ids the opened document already spent, which a new comment may not take */
export function reservedCommentIds(state: EditorState): ReadonlySet<string> {
  return documentOf(state).reservedCommentIds;
}

/** The paragraph ids the opened comment parts already spent */
export function reservedCommentParaIds(
  state: EditorState
): ReadonlySet<string> {
  return documentOf(state).reservedCommentParaIds;
}
