/**
 * Brings what the document laid down into the editor: what styles.xml told us, and the paper it
 * is written on.
 *
 * Commands and the toolbar know nothing about the session, so this plugin holds all of it on
 * their behalf: deriving the display values again after a formatting edit needs the values the
 * style laid down, the toolbar needs the size a run is rendered at where nobody wrote a size
 * down, and a new table is fitted to the width of the body the paper leaves.
 * The values are fixed the moment the document is opened and do not change while editing.
 */

import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { DEFAULT_TAB_STOP_PT } from "../docx/documentSettings";
import {
  NO_DOCUMENT_DEFAULTS,
  NO_PARAGRAPH_FORMATTING,
  NO_STYLES,
  type ParagraphFormatLayer,
  type ParagraphFormattingContext,
  type ParagraphStyleOption,
  type StyleTable,
} from "../docx/formatting";
import {
  A4_PORTRAIT,
  bodyHeightTwips,
  bodyWidth,
  type PageGeometry,
  twipsToPx,
} from "../docx/pageGeometry";
import type { DocumentDefaults } from "../model/format";
import { EMPTY_NUMBERING, type Numbering } from "../numbering/parseNumbering";

/** Values that apply to the whole document */
interface DocumentInfo {
  styles: StyleTable;
  defaults: DocumentDefaults;
  paragraphDefaults: ParagraphFormatLayer;
  numbering: Numbering;
  paragraphStyles: ParagraphStyleOption[];
  /** The paper the document is written on, which decides the width a new table is fitted to */
  geometry: PageGeometry;
  /** The distance between automatic tab stops when no custom stop applies. */
  defaultTabStopPt: number;
}

/** A document whose paragraph styles the editor does not know offers none to pick from */
const NO_PARAGRAPH_STYLES: ParagraphStyleOption[] = [];

const documentKey = new PluginKey<DocumentInfo>("docxEditorDocument");

/** This document's style table. Empty when the editor does not know the table */
export function documentStyles(state: EditorState): StyleTable {
  return documentKey.getState(state)?.styles ?? NO_STYLES;
}

/** The default formatting this document wrote down. When the editor does not know it, it is the same as nothing being specified */
export function documentDefaults(state: EditorState): DocumentDefaults {
  return documentKey.getState(state)?.defaults ?? NO_DOCUMENT_DEFAULTS;
}

/** The values required to resolve the OOXML paragraph-property hierarchy. */
export function documentParagraphFormatting(
  state: EditorState
): ParagraphFormattingContext {
  const info = documentKey.getState(state);
  if (!info) return NO_PARAGRAPH_FORMATTING;
  return {
    styles: info.styles,
    defaultStyleId:
      info.paragraphStyles.find((style) => style.isDefault)?.id ?? null,
    defaults: info.paragraphDefaults,
    numbering: info.numbering,
  };
}

/** The paragraph styles this document defines. Empty when the editor does not know them */
export function documentParagraphStyles(
  state: EditorState
): ParagraphStyleOption[] {
  return documentKey.getState(state)?.paragraphStyles ?? NO_PARAGRAPH_STYLES;
}

/**
 * The style a paragraph pointing at none of its own wears (`w:default="1"`). Null when the
 * document marks none.
 *
 * It is the style the list of them marks, of which `readParagraphStyles` leaves exactly one.
 */
export function defaultParagraphStyleId(state: EditorState): string | null {
  return (
    documentParagraphStyles(state).find((style) => style.isDefault)?.id ?? null
  );
}

/**
 * The paper the document is written on. A document the editor knows no paper for is drawn on
 * the same A4 every document was drawn on before the geometry was read
 */
export function documentGeometry(state: EditorState): PageGeometry {
  return documentKey.getState(state)?.geometry ?? A4_PORTRAIT;
}

/** The document-wide interval between automatic tab stops, in points. */
export function documentDefaultTabStopPt(state: EditorState): number {
  return documentKey.getState(state)?.defaultTabStopPt ?? DEFAULT_TAB_STOP_PT;
}

/**
 * The width one line of body text occupies on this document's paper, in pixels.
 *
 * It is the width an image wider than the page is shrunk to, so a caller inserting an image
 * itself asks for it rather than taking the A4 width `fittedExtent` falls back to.
 */
export function documentBodyWidthPx(state: EditorState): number {
  return bodyWidth(documentGeometry(state)).px;
}

/** The height one page of body content occupies, in pixels. */
export function documentBodyHeightPx(state: EditorState): number {
  return twipsToPx(bodyHeightTwips(documentGeometry(state)));
}

export function documentStyleTable(
  styles: StyleTable,
  defaults: DocumentDefaults = NO_DOCUMENT_DEFAULTS,
  paragraphStyles: ParagraphStyleOption[] = NO_PARAGRAPH_STYLES,
  geometry: PageGeometry = A4_PORTRAIT,
  paragraphDefaults: ParagraphFormatLayer = {},
  numbering: Numbering = EMPTY_NUMBERING,
  defaultTabStopPt: number = DEFAULT_TAB_STOP_PT
): Plugin<DocumentInfo> {
  return new Plugin<DocumentInfo>({
    key: documentKey,
    state: {
      init: () => ({
        styles,
        defaults,
        paragraphDefaults,
        numbering,
        paragraphStyles,
        geometry,
        defaultTabStopPt,
      }),
      apply: (_tr, current) => current,
    },
  });
}
