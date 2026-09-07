/**
 * What the opened document laid down, held as one value: the style table and the list
 * definitions, the paper the body is written on, the paragraph defaults every value is resolved
 * against, and the identifiers its comment parts have already spent.
 *
 * Commands and the toolbar know nothing about the session, so the snapshot stands for it on their
 * behalf, and `editorDocumentOf` is the one place a session is read into editor values: a screen
 * and a test that open the same document hold the same values.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { DEFAULT_TAB_STOP_PT } from "../docx/documentSettings";
import {
  NO_DOCUMENT_DEFAULTS,
  NO_STYLES,
  type ParagraphFormatLayer,
  type ParagraphStyleOption,
  type StyleTable,
} from "../docx/formatting";
import { A4_PORTRAIT, type PageGeometry } from "../docx/pageGeometry";
import {
  type SessionStore,
  documentNumbering as sessionNumbering,
} from "../docx/session";
import type { DocumentDefaults } from "../model/format";
import { EMPTY_NUMBERING, type Numbering } from "../numbering/parseNumbering";

/** The document-level values one editing state is built on */
export interface EditorDocument {
  /** The document these values were read from. Null for a state built without one */
  readonly session: SessionStore | null;
  readonly styles: StyleTable;
  readonly defaults: DocumentDefaults;
  /** Paragraph properties from the styles.xml docDefaults */
  readonly paragraphDefaults: ParagraphFormatLayer;
  /** The styles the document defines for the style picker to offer */
  readonly paragraphStyles: ParagraphStyleOption[];
  /**
   * The style a paragraph pointing at none of its own wears (`w:default="1"`). Null when the
   * document marks none
   */
  readonly defaultParagraphStyleId: string | null;
  readonly numbering: Numbering;
  /**
   * Whether the document has a place (numbering.xml) to write the definition of a new list.
   * A state built without opening a document assumes that place exists and behaves as every
   * state did before the question was asked.
   */
  readonly canStartNewList: boolean;
  /** The paper the document is written on, which decides the width a new table is fitted to */
  readonly geometry: PageGeometry;
  /** The distance between automatic tab stops when no custom stop applies */
  readonly defaultTabStopPt: number;
  /** Every id already present in the opened Comments part, including unreferenced entries */
  readonly reservedCommentIds: ReadonlySet<string>;
  /** Every paragraph id present in the opened comment parts, including orphan extension entries */
  readonly reservedCommentParaIds: ReadonlySet<string>;
}

const NO_IDS: ReadonlySet<string> = new Set();

/** A document whose paragraph styles the editor does not know offers none to pick from */
const NO_PARAGRAPH_STYLES: ParagraphStyleOption[] = [];

/** What a state built without an opened document answers: the same as nothing being written down */
export const NO_DOCUMENT: EditorDocument = {
  session: null,
  styles: NO_STYLES,
  defaults: NO_DOCUMENT_DEFAULTS,
  paragraphDefaults: {},
  paragraphStyles: NO_PARAGRAPH_STYLES,
  defaultParagraphStyleId: null,
  numbering: EMPTY_NUMBERING,
  canStartNewList: true,
  geometry: A4_PORTRAIT,
  defaultTabStopPt: DEFAULT_TAB_STOP_PT,
  reservedCommentIds: NO_IDS,
  reservedCommentParaIds: NO_IDS,
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

/** The one place an opened document is read into the values the editor holds */
export function editorDocumentOf(session: SessionStore): EditorDocument {
  return {
    session,
    styles: session.styles,
    defaults: session.defaults,
    paragraphDefaults: session.paragraphDefaults,
    paragraphStyles: session.paragraphStyles,
    defaultParagraphStyleId: session.defaultParagraphStyleId,
    numbering: sessionNumbering(session),
    canStartNewList: session.numberingPartPath !== null,
    geometry: session.geometry,
    defaultTabStopPt: session.defaultTabStopPt,
    reservedCommentIds: new Set(session.comments.byId.keys()),
    reservedCommentParaIds: reservedParaIds(session),
  };
}

const editorDocumentKey = new PluginKey<EditorDocument>("docxEditorDocument");

/** The document the state was built for. `NO_DOCUMENT` for a state built without one */
export function documentOf(state: EditorState): EditorDocument {
  return editorDocumentKey.getState(state) ?? NO_DOCUMENT;
}

/**
 * The snapshot for an opened document held against the document node it was built over.
 *
 * Today every value is the session's and the document node carries none of them, so `_doc` is
 * read for nothing. A document-level edit is recorded on the document node's attrs, which is what
 * the export reads and what undo carries, and this is where the snapshot follows it.
 */
function derive(session: SessionStore, _doc: PMNode): EditorDocument {
  return editorDocumentOf(session);
}

/**
 * Holds the snapshot for the lifetime of the state.
 *
 * It is derived again only where a document-level edit could have been recorded, so an ordinary
 * edit and a selection move both leave the same object behind, allowing the sheet's style to be
 * cached against that identity.
 */
export function editorDocument(
  document: EditorDocument
): Plugin<EditorDocument> {
  return new Plugin<EditorDocument>({
    key: editorDocumentKey,
    state: {
      init: () => document,
      apply: (tr, current, old) =>
        current.session === null || tr.doc.attrs === old.doc.attrs
          ? current
          : derive(current.session, tr.doc),
    },
  });
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
