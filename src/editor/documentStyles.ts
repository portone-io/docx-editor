/**
 * What the document laid down, as the commands and the toolbar ask for it: what styles.xml told
 * us, and the paper it is written on.
 *
 * Commands and the toolbar know nothing about the session, so they read the snapshot
 * `editor/editorDocument` holds on their behalf: deriving the display values again after a
 * formatting edit needs the values the style laid down, the toolbar needs the size a run is
 * rendered at where nobody wrote a size down, and a new table is fitted to the width of the body
 * the paper leaves.
 */

import type { EditorState } from "prosemirror-state";
import type {
  ParagraphFormattingContext,
  ParagraphStyleOption,
  StyleTable,
} from "../docx/formatting";
import {
  bodyHeightTwips,
  bodyWidth,
  type PageGeometry,
  twipsToPx,
} from "../docx/pageGeometry";
import type { DocumentDefaults } from "../model/format";
import { documentOf } from "./editorDocument";

/** This document's style table. Empty when the editor does not know the table */
export function documentStyles(state: EditorState): StyleTable {
  return documentOf(state).styles;
}

/** The default formatting this document wrote down. When the editor does not know it, it is the same as nothing being specified */
export function documentDefaults(state: EditorState): DocumentDefaults {
  return documentOf(state).defaults;
}

/** The values required to resolve the OOXML paragraph-property hierarchy. */
export function documentParagraphFormatting(
  state: EditorState
): ParagraphFormattingContext {
  const document = documentOf(state);
  return {
    styles: document.styles,
    defaultStyleId: document.defaultParagraphStyleId,
    defaults: document.paragraphDefaults,
    numbering: document.numbering,
  };
}

/** The paragraph styles this document defines. Empty when the editor does not know them */
export function documentParagraphStyles(
  state: EditorState
): ParagraphStyleOption[] {
  return documentOf(state).paragraphStyles;
}

/**
 * The style a paragraph pointing at none of its own wears (`w:default="1"`). Null when the
 * document marks none.
 */
export function defaultParagraphStyleId(state: EditorState): string | null {
  return documentOf(state).defaultParagraphStyleId;
}

/**
 * The paper the document is written on. A document the editor knows no paper for is drawn on
 * the same A4 every document was drawn on before the geometry was read
 */
export function documentGeometry(state: EditorState): PageGeometry {
  return documentOf(state).geometry;
}

/** The document-wide interval between automatic tab stops, in points. */
export function documentDefaultTabStopPt(state: EditorState): number {
  return documentOf(state).defaultTabStopPt;
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
