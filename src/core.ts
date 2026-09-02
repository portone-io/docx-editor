/**
 * Framework-free DOCX import and export API. The opaque session preserves package parts between
 * calls; server runtimes must provide global `DOMParser` and `Node` implementations.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type DocxBytes, importDocx as openDocx } from "./docx/importDocx";
import type { DocxSession } from "./docx/session";
import { commentAdditionsBy, protectionAllows } from "./schema/protection";

export { exportDocx } from "./docx/exportDocx";
export type { ParagraphStyleOption } from "./docx/formatting";
export type { DocxBytes } from "./docx/importDocx";
export type { DocxSession } from "./docx/session";
export { documentNumbering, documentPartPath } from "./docx/session";
export type {
  CellFormat,
  CellVerticalAlign,
  DocumentDefaults,
  HighlightName,
  LineSpacing,
  NumberingRef,
  ParagraphAlign,
  ParagraphFormat,
  RowFormat,
  RowHeight,
  RunFormat,
  TableFormat,
  TableWidth,
  TableWidthType,
  UnderlineKind,
  VerticalAlign,
} from "./model/format";
export {
  toCellFormat,
  toParagraphFormat,
  toRowFormat,
  toRunFormat,
  toTableFormat,
  toTableWidth,
} from "./model/format";
export type {
  LevelIndent,
  NumberFormat,
  Numbering,
  NumberingLevel,
  NumberingList,
} from "./numbering/parseNumbering";
export { parseNumbering } from "./numbering/parseNumbering";
export type {
  DocxExportErrorCode,
  DocxImportErrorCode,
} from "./ooxml/errors";
export { DocxExportError, DocxImportError } from "./ooxml/errors";
/**
 * An image node's size is written in EMU, the unit the document itself uses. A builder
 * that places an image is working in pixels, so the conversion ships with the type.
 */
export type { ImageExtent } from "./ooxml/image";
export { emuToPx, pxToEmu, toImageExtent } from "./ooxml/image";
export { docxSchema } from "./schema";

/**
 * Opens docx bytes into the document to work on and the session that remembers the file they came
 * from.
 *
 * The engine hands out the store it fills in; this is where it narrows to the opaque session, so
 * that the original XML behind it stays the exporter's business.
 */
export function importDocx(input: DocxBytes): {
  doc: PMNode;
  session: DocxSession;
} {
  return openDocx(input);
}

/**
 * Whether the second document differs from the first in nothing but comments, every one of them
 * added, edited, deleted, replied to or settled by the author with this identity.
 *
 * This is the judgement the editor makes under `mode: { kind: "comment" }`, made again over the
 * file a browser handed back: the editor's own refusal is a courtesy to the user, and a server
 * taking the file in is where the rule holds. A comment carrying no recorded identity is
 * everyone's to edit here as it is in the editor (`schema/protection`); a comment that appeared
 * has to carry this identity, since a file can claim any author and the editor's own hand in
 * writing it is not there to vouch for it.
 */
export function onlyCommentsChangedBy(
  original: DocxBytes,
  submitted: DocxBytes,
  authorId: string
): boolean {
  const before = openDocx(original).doc;
  const after = openDocx(submitted).doc;
  return (
    protectionAllows(before, after, {
      protection: "comments",
      authorId,
      editableComments: "own",
    }) && commentAdditionsBy(before, after, authorId)
  );
}
