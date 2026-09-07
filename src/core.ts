/**
 * Framework-free DOCX import and export API. The opaque session preserves package parts between
 * calls.
 *
 * Reading a package needs an XML parser. A browser has one; anywhere else, hand one in as
 * `xmlParser` or install a `DOMParser` global, or the call is refused with the import code
 * `no-xml-parser`.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { FidelityNote } from "./docx/fidelity";
import {
  type DocxBytes,
  type ImportOptions,
  importDocx as openDocx,
} from "./docx/importDocx";
import type { DocxSession } from "./docx/session";

export type { CommentOnlyVerdict } from "./docx/commentOnlyChange";
export { onlyCommentsChangedBy } from "./docx/commentOnlyChange";
export type { ExportOptions } from "./docx/exportDocx";
export { exportDocx, exportDocxReport } from "./docx/exportDocx";
export type {
  FidelityCode,
  FidelityNote,
  FidelitySeverity,
} from "./docx/fidelity";
export type { ParagraphStyleOption } from "./docx/formatting";
export type { DocxBytes, ImportOptions } from "./docx/importDocx";
/**
 * What the writer would refuse the document over, asked ahead of the write. The first entry is
 * what `exportDocx` throws, so the two cannot disagree.
 */
export type { ExportProblem } from "./docx/invariants";
export { exportProblems } from "./docx/invariants";
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
  NumberingOptions,
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
export type { XmlParser } from "./ooxml/xml";
export { docxSchema } from "./schema";

/**
 * Opens docx bytes into the document to work on, the session that remembers the file they came
 * from, and what the file holds that this editor could not model.
 *
 * The engine hands out the store it fills in; this is where it narrows to the opaque session, so
 * that the original XML behind it stays the exporter's business.
 */
export function importDocx(
  input: DocxBytes,
  options?: ImportOptions
): {
  doc: PMNode;
  session: DocxSession;
  notes: FidelityNote[];
} {
  return openDocx(input, options);
}
