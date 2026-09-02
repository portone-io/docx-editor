/**
 * This file knows nothing about zip or XML. All it knows is the document model and the screen.
 */

import { baseKeymap } from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { history } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Plugin } from "prosemirror-state";
import { tableEditing } from "prosemirror-tables";
import { EditorView } from "prosemirror-view";
import { DEFAULT_TAB_STOP_PT } from "../docx/documentSettings";
import {
  NO_DOCUMENT_DEFAULTS,
  NO_STYLES,
  type ParagraphFormatLayer,
  type ParagraphStyleOption,
  type StyleTable,
} from "../docx/formatting";
import { A4_PORTRAIT, type PageGeometry } from "../docx/pageGeometry";
import type { DocumentDefaults } from "../model/format";
import { EMPTY_NUMBERING, type Numbering } from "../numbering/parseNumbering";
import { pageDecorations } from "../page/pageDecorations";
import { pageGeometryStyle, pagePixels } from "../page/pageLayout";
import type { EditableComments, EditingProtection } from "../schema/protection";
import { editsShut } from "../schema/protectionState";
import { editorClassNames } from "../styles/classNames";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../styles/fontStack";
import { documentDefaultsStyle } from "../styles/inlineStyle";
import { gridBorders, withDerivedGridBorders } from "../table/gridBorders";
import type { CommentAuthor } from "./commands/comments/model";
import { documentDefaultTabStopPt, documentStyleTable } from "./documentStyles";
import { externalClipboard } from "./externalClipboard";
import { imageFiles } from "./imageFiles";
import { bookmarkProtection } from "./plugins/bookmarkProtection";
import { columnResize } from "./plugins/columnResize";
import { commentDecorations } from "./plugins/commentDecorations";
import { commentReservations } from "./plugins/commentReservations";
import { documentProtection } from "./plugins/documentProtection";
import { imagePaste } from "./plugins/imagePaste";
import { docxKeymap, historyKeys } from "./plugins/keymap";
import { linkPanel } from "./plugins/linkPanel";
import { listInputRules } from "./plugins/listInputRules";
import { lockedContent } from "./plugins/lockedContent";
import { noteProtection } from "./plugins/noteProtection";
import { numberingMarkers } from "./plugins/numberingDecorations";
import { rowResize } from "./plugins/rowResize";
import { styledParagraphs } from "./plugins/styledParagraphs";
import { tabCaret } from "./plugins/tabCaret";
import { tabDecorations } from "./plugins/tabDecorations";
import { tabLayout } from "./plugins/tabLayout";
import { tableContextMenu } from "./plugins/tableContextMenu";
import { tabPointer } from "./plugins/tabPointer";
import { textContextMenu } from "./plugins/textContextMenu";
import { ImageNodeView } from "./views/imageResize";
import { runMarkView } from "./views/runMarkView";

export interface EditorStateOptions {
  numbering?: Numbering;
  styles?: StyleTable;
  defaults?: DocumentDefaults;
  /** Paragraph properties from styles.xml docDefaults. */
  paragraphDefaults?: ParagraphFormatLayer;
  /**
   * Whether the document has a place (numbering.xml) to write the definition of a new list.
   * Callers that build a state without opening a document assume that place exists and behave
   * as they do today.
   */
  canStartNewList?: boolean;
  /** The plugins handed in from outside the package */
  consumerPlugins?: readonly Plugin[];
  /** The styles the document defines for the style picker to offer */
  paragraphStyles?: ParagraphStyleOption[];
  /**
   * Whether the right click is the editor's own. Turned off, the browser's own menu is never
   * taken away, which is what a consumer drawing menus of its own needs.
   */
  contextMenus?: boolean;
  /** The paper the document names. A4 where a document names none */
  geometry?: PageGeometry;
  /** The interval between automatic tab stops, in points. */
  defaultTabStopPt?: number;
  /** Every id already present in the opened Comments part, including unreferenced entries. */
  reservedCommentIds?: Iterable<string>;
  /** Every paragraph id present in the opened comment parts, including orphan extension entries. */
  reservedCommentParaIds?: Iterable<string>;
  /**
   * What the document as a whole may receive (`schema/protection`): everything, comments alone,
   * or nothing. Everything when none is given, which is what every state was before.
   */
  protection?: EditingProtection;
  /** Whose comments are written, and so whose may be edited under `editableComments: "own"` */
  author?: CommentAuthor | null;
  /** Whose comments may be edited or deleted. One's own when none is given */
  editableComments?: EditableComments;
}

/**
 * Creates a single editing state.
 * The list definitions, the style table, and the document defaults are what the document
 * wrote down, and commands and the toolbar read them through plugins.
 */
export function createEditorState(
  doc: PMNode,
  {
    numbering = EMPTY_NUMBERING,
    styles = NO_STYLES,
    defaults = NO_DOCUMENT_DEFAULTS,
    paragraphDefaults = {},
    canStartNewList = true,
    consumerPlugins = [],
    paragraphStyles = [],
    contextMenus = true,
    geometry = A4_PORTRAIT,
    defaultTabStopPt = DEFAULT_TAB_STOP_PT,
    reservedCommentIds = [],
    reservedCommentParaIds = [],
    protection = "none",
    author = null,
    editableComments = "own",
  }: EditorStateOptions = {}
): EditorState {
  return EditorState.create({
    doc: withDerivedGridBorders(doc),
    plugins: [
      // Consumer plugins lead the array. ProseMirror walks the plugins in order and takes the
      // first answer for a keypress, a paste, a drop or any other DOM event, so this is the
      // only place from which a consumer handler can win over the built-in one.
      ...consumerPlugins,
      // Refuses every edit inside a locked content control, whoever asked for it. It is not
      // optional: a document that locked a part of itself stays locked in every consumer.
      // The same guard refuses what the protection below shuts.
      lockedContent(),
      documentProtection({ protection, author, editableComments }),
      // Bookmark ranges are preserved rather than edited, so neither half may disappear.
      bookmarkProtection(),
      // Note bodies are display-only, so their imported main-story references stay paired.
      noteProtection(),
      // An orphan Comments-part entry still owns its id and must not be replaced by a new comment.
      commentReservations(reservedCommentIds, reservedCommentParaIds),
      history(),
      keymap(docxKeymap),
      historyKeys(),
      keymap(baseKeymap),
      // Holds whether the link panel is open, which Cmd+K above and the toolbar button both set
      linkPanel(),
      // Turns a typed list prefix ("1. ", "- ") into a list. Backspace, bound in `docxKeymap`
      // above, is what takes such a conversion back
      listInputRules(),
      dropCursor(),
      // Clipboard priority is files, resolvable HTML images, then regular HTML or plain text.
      imageFiles(),
      imagePaste(),
      externalClipboard(),
      // A press that grabs a table edge must be intercepted before a cell-selection drag starts.
      // For DOM events the plugin registered first wins, so both resizers precede `tableEditing`.
      columnResize(),
      rowResize(),
      tableEditing(),
      // Comment anchors live in the model; this paints their ranges without changing document XML.
      commentDecorations(),
      // Adjacent text tabs still need separate DOM ranges for layout and pointer selection.
      tabDecorations(),
      tabPointer(),
      tabLayout(),
      tabCaret(),
      // Keeps table-cell lines aligned with OOXML precedence after an edit
      gridBorders(),
      // Reads the document's styles into the paragraphs an edit built from nothing
      styledParagraphs(),
      // Receiving a right-click means taking the browser's own menu away, so both of these stand
      // or fall together with the menus the editor draws. The text menu stands ahead of the table
      // menu, and hands a click with nothing selected inside a cell back to it
      ...(contextMenus ? [textContextMenu(), tableContextMenu()] : []),
      numberingMarkers(numbering, canStartNewList),
      // Values only go in when page display is turned on
      pageDecorations(),
      documentStyleTable(
        styles,
        defaults,
        paragraphStyles,
        geometry,
        paragraphDefaults,
        numbering,
        defaultTabStopPt
      ),
    ],
  });
}

export interface EditorOptions {
  mount: HTMLElement;
  state: EditorState;
  defaults: DocumentDefaults;
  /** The paper the document names. The sheet is drawn from it, A4 where a document names none */
  geometry?: PageGeometry;
  readOnly: boolean;
  /** The fonts stood in for the ones the document declares. The built-in set when none is given */
  fontFallbacks?: FontFallbacks;
  onStateChange: (state: EditorState) => void;
}

export function createEditorView({
  mount,
  state,
  defaults,
  geometry = A4_PORTRAIT,
  readOnly,
  fontFallbacks = DEFAULT_FONT_FALLBACKS,
  onStateChange,
}: EditorOptions): EditorView {
  const view = new EditorView(mount, {
    state,
    // A protection that shuts the body shuts typing with it. Selecting stays open either way, which
    // is what a reader marking a stretch for a comment needs
    editable: (current) => !readOnly && !editsShut(current),
    attributes: {
      class: editorClassNames.sheet,
      // The paper first, so a document that names one is drawn on it from the first frame
      style:
        `${pageGeometryStyle(pagePixels(geometry))};` +
        `${documentDefaultsStyle(defaults, fontFallbacks)};` +
        `tab-size:${documentDefaultTabStopPt(state)}pt`,
    },
    // The schema can only draw a run with the default fallback fonts, so this editor draws its own
    markViews: { run: runMarkView(fontFallbacks) },
    // An image is drawn by a view of its own, which is what carries the resize handles
    nodeViews: {
      image: (node, imageView, getPos) =>
        new ImageNodeView(node, imageView, getPos),
    },
    dispatchTransaction(transaction) {
      const next = view.state.apply(transaction);
      view.updateState(next);
      onStateChange(next);
    },
  });
  return view;
}
