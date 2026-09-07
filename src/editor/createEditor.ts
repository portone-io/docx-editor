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
import type { SessionStore } from "../docx/session";
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
import {
  documentOf,
  type EditorDocument,
  editorDocument,
  editorDocumentOf,
  NO_DOCUMENT,
} from "./editorDocument";
import { externalClipboard } from "./externalClipboard";
import { imageFiles } from "./imageFiles";
import { columnResize } from "./plugins/columnResize";
import { commentDecorations } from "./plugins/commentDecorations";
import { documentProtection } from "./plugins/documentProtection";
import { imagePaste } from "./plugins/imagePaste";
import { docxKeymap, historyKeys } from "./plugins/keymap";
import { linkPanel } from "./plugins/linkPanel";
import { listInputRules } from "./plugins/listInputRules";
import { lockedContent } from "./plugins/lockedContent";
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
  /**
   * What the opened document laid down: its styles, its list definitions, the paper it is
   * written on. A state built without one reads `NO_DOCUMENT`, which answers as nothing having
   * been written down. `editorStateForSession` is what fills it in for an opened document.
   */
  document?: EditorDocument;
  /** The plugins handed in from outside the package */
  consumerPlugins?: readonly Plugin[];
  /**
   * Whether the right click is the editor's own. Turned off, the browser's own menu is never
   * taken away, which is what a consumer drawing menus of its own needs.
   */
  contextMenus?: boolean;
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
 * wrote down, and commands and the toolbar read them off the snapshot the state holds.
 */
export function createEditorState(
  doc: PMNode,
  options: EditorStateOptions = {}
): EditorState {
  const {
    document = NO_DOCUMENT,
    consumerPlugins = [],
    contextMenus = true,
    protection = "none",
    author = null,
    editableComments = "own",
  } = options;
  return EditorState.create({
    doc: withDerivedGridBorders(doc),
    plugins: [
      // Consumer plugins lead the array. ProseMirror walks the plugins in order and takes the
      // first answer for a keypress, a paste, a drop or any other DOM event, so this is the
      // only place from which a consumer handler can win over the built-in one.
      ...consumerPlugins,
      // Every document-level value the editor reads stands in one snapshot, and it leads the
      // built-in plugins so that the ones drawing from it are initialized after it
      editorDocument(document),
      // Refuses every edit no guard in `schema/guards` lets through, whoever asked for it. It is
      // not optional: a document that locked a part of itself stays locked in every consumer, and
      // the preserved bookmark markers and note references stay where the file put them.
      lockedContent(),
      documentProtection({ protection, author, editableComments }),
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
      numberingMarkers(),
      // Values only go in when page display is turned on
      pageDecorations(),
    ],
  });
}

/**
 * The state for a document that was opened, which is the one way a session becomes an editing
 * state: what the screen builds and what a test builds are then the same values.
 */
export function editorStateForSession(
  opened: { doc: PMNode; session: SessionStore },
  options: Omit<EditorStateOptions, "document"> = {}
): EditorState {
  return createEditorState(opened.doc, {
    ...options,
    document: editorDocumentOf(opened.session),
  });
}

/**
 * The sheet one document is drawn on: the paper it names and the defaults it wrote down.
 *
 * Both come off the state, so a document-level edit reaches the sheet, and both are remembered
 * against the snapshot they were built from, which an ordinary edit leaves as it is.
 */
const sheetStyles = new WeakMap<
  EditorDocument,
  { fallbacks: FontFallbacks; style: string }
>();

function sheetStyleOf(
  document: EditorDocument,
  fallbacks: FontFallbacks
): string {
  const remembered = sheetStyles.get(document);
  if (remembered?.fallbacks === fallbacks) return remembered.style;
  const style =
    `${pageGeometryStyle(pagePixels(document.geometry))};` +
    `${documentDefaultsStyle(document.defaults, fallbacks)};`;
  sheetStyles.set(document, { fallbacks, style });
  return style;
}

export interface EditorOptions {
  mount: HTMLElement;
  state: EditorState;
  /** The fonts stood in for the ones the document declares. The built-in set when none is given */
  fontFallbacks?: FontFallbacks;
  onStateChange: (state: EditorState) => void;
}

export function createEditorView({
  mount,
  state,
  fontFallbacks = DEFAULT_FONT_FALLBACKS,
  onStateChange,
}: EditorOptions): EditorView {
  const view = new EditorView(mount, {
    state,
    // A protection that shuts the body shuts typing with it, and is read off the state so that a
    // mode switched on an open document takes effect without a new view. Selecting stays open
    // either way, which is what a reader marking a stretch for a comment needs
    editable: (current) => !editsShut(current),
    attributes: (current) => {
      const document = documentOf(current);
      return {
        class: editorClassNames.sheet,
        // The paper first, so a document that names one is drawn on it from the first frame
        style: `${sheetStyleOf(document, fontFallbacks)}tab-size:${document.defaultTabStopPt}pt`,
        // A sheet that takes no typing is no longer focusable of itself, so a reader or a
        // commenter is handed the focus another way: the keys reach it, and the selection stays
        // its own
        ...(editsShut(current) ? { tabindex: "0" } : {}),
      };
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
