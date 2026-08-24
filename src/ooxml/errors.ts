/**
 * The error kinds used when opening and exporting a docx.
 *
 * Messages are English one-liners meant for developers. The `code` is the stable,
 * machine-readable part: consumers switch on it to show their own localized text,
 * so a code never changes meaning even when a message is reworded.
 */

/**
 * Why a document could not be opened.
 *
 * - `not-a-docx`: the bytes are not a readable zip container, its entry names are not a package's, or an entry does not hold what it says it does
 * - `too-large`: the package asks to inflate to more than we open
 * - `missing-part`: the package has no main document part to read
 * - `missing-body`: the main part carries no `w:body`
 * - `malformed-xml`: the XML cannot be parsed, declares a DTD, or its markup is inconsistent
 * - `unsupported-content`: the document holds markup we cannot write back out unchanged
 */
export type DocxImportErrorCode =
  | "not-a-docx"
  | "too-large"
  | "missing-part"
  | "missing-body"
  | "malformed-xml"
  | "unsupported-content";

/** Thrown when we hit a document whose content cannot be kept safely. We refuse to open it instead of losing it silently */
export class DocxImportError extends Error {
  readonly code: DocxImportErrorCode;

  constructor(
    code: DocxImportErrorCode,
    message: string,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "DocxImportError";
    this.code = code;
  }
}

/**
 * Why a document could not be written back out.
 *
 * - `missing-numbering-part`: a new list needs a numbering.xml the document does not have
 * - `missing-content-types`: a new image needs a [Content_Types].xml the package does not have
 * - `unsupported-content`: the document holds a node kind we have no way to serialize
 * - `lost-original`: a node that only carries its original XML has lost it
 * - `malformed-xml`: an original XML fragment cannot be read well enough to rewrite
 * - `invalid-table`: the table grid is inconsistent, e.g. a vertical merge outliving its rows
 */
export type DocxExportErrorCode =
  | "missing-numbering-part"
  | "missing-content-types"
  | "unsupported-content"
  | "lost-original"
  | "malformed-xml"
  | "invalid-table";

/** Thrown when an edited document cannot be written back out without losing or corrupting content */
export class DocxExportError extends Error {
  readonly code: DocxExportErrorCode;

  constructor(
    code: DocxExportErrorCode,
    message: string,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = "DocxExportError";
    this.code = code;
  }
}
