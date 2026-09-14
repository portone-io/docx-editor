/**
 * The error kinds used when opening and exporting a docx.
 *
 * Messages are English one-liners meant for developers. The `code` is the stable,
 * machine-readable part: consumers switch on it to show their own localized text,
 * so a code never changes meaning even when a message is reworded. An export refusal carries an
 * `ExportProblemReason` besides, which tells the several situations one code covers apart and
 * names what each is about, so a host can say which content to put back.
 */

/**
 * Why a document could not be opened.
 *
 * - `no-xml-parser`: the runtime has no XML parser to read the package with. The one code here that is about where the call was made rather than about the file
 * - `not-a-docx`: the bytes are not a readable zip container, its entry names are not a package's, or an entry does not hold what it says it does
 * - `too-large`: the package asks to inflate to more than we open
 * - `missing-part`: the package has no main document part to read
 * - `missing-body`: the main part carries no `w:body`
 * - `malformed-xml`: the XML cannot be parsed, declares a DTD, or its markup is inconsistent
 * - `unsupported-conformance`: the package is an ECMA-376 Strict one, and this editor reads Transitional
 * - `unsupported-content`: the document holds markup we cannot write back out unchanged
 */
export type DocxImportErrorCode =
  | "no-xml-parser"
  | "not-a-docx"
  | "too-large"
  | "missing-part"
  | "missing-body"
  | "malformed-xml"
  | "unsupported-conformance"
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
 * - `missing-content-types`: a part the writer adds, be it an image, a list definition or a comment, needs a [Content_Types].xml the package does not have
 * - `unsupported-content`: the document holds something no correct file can be written from: a node kind we have no way to serialize, a preserved block standing in two places, which has one original XML to write, or a paragraph in a list nothing defines
 * - `lost-original`: a node that only carries its original XML has lost it
 * - `malformed-xml`: an original XML fragment cannot be read well enough to rewrite
 * - `invalid-table`: the table grid is inconsistent, e.g. a vertical merge outliving its rows
 */
export type DocxExportErrorCode =
  | "missing-content-types"
  | "unsupported-content"
  | "lost-original"
  | "malformed-xml"
  | "invalid-table";

/**
 * The kinds of side story a refusal can be about, which are the kinds a document holds beside its
 * body (`schema/stories`). Spelled here because the reasons stand below the schema layer.
 */
export type ExportStoryKind =
  | "comment"
  | "footnote"
  | "endnote"
  | "header"
  | "footer";

/** One side story, as the document holds it */
export interface ExportProblemStory {
  readonly kind: ExportStoryKind;
  /** A note's number, a comment's id, or the path of the part a header or footer stands in */
  readonly id: string;
}

/** A part the export writes beside the body, named after what it holds */
export type ExportPartName =
  | "media"
  | "numbering"
  | "comments"
  | "commentsExtended"
  | "footnotes"
  | "endnotes";

/**
 * What a refusal is actually about.
 *
 * A `code` groups several situations - `unsupported-content` alone covers a list nothing defines,
 * a block standing twice, and a change to a story no writer carries - and the `message` telling
 * them apart is an English one-liner for a developer. This says which situation was met and names
 * what it is about, so a host can write the sentence its own user reads. Switch on `kind`.
 */
export type ExportProblemReason =
  /** A preserved fragment holding bookmark markup could not be parsed (`malformed-xml`) */
  | { readonly kind: "unreadable-preserved-xml" }
  /** A bookmark marker carries no id to pair it by (`malformed-xml`) */
  | { readonly kind: "unnamed-bookmark"; readonly marker: "start" | "end" }
  /**
   * A bookmark marker whose partner is gone (`malformed-xml`): a `start` with no end after it, or
   * an `end` with no start before it.
   */
  | {
      readonly kind: "unmatched-bookmark";
      readonly id: string;
      readonly marker: "start" | "end";
    }
  /** One bookmark id is started twice (`malformed-xml`) */
  | { readonly kind: "repeated-bookmark-start"; readonly id: string }
  /** A part cannot be rewritten around its root element (`malformed-xml`) */
  | {
      readonly kind: "unwritable-part-root";
      readonly part: ExportPartName;
    }
  /** A cell's vertical merge covers rows its table does not have (`invalid-table`) */
  | { readonly kind: "vertical-merge-past-table" }
  /** A node that is written from its original XML alone no longer holds it (`lost-original`) */
  | { readonly kind: "lost-preserved-xml"; readonly node: string }
  /** A placeholder pasted in from another opened document, whose XML this session never read (`lost-original`) */
  | {
      readonly kind: "preserved-from-another-document";
      readonly node: string;
      /** The session the placeholder was opened in */
      readonly sessionId: string;
    }
  /** One preserved block stands in two places, and it has one original XML to be written (`unsupported-content`) */
  | {
      readonly kind: "duplicate-preserved-block";
      readonly node: string;
      /** The side story it stands in, and null for a block of the body */
      readonly story: ExportProblemStory | null;
    }
  /** A paragraph is in a list neither the file nor the editor's register defines (`unsupported-content`) */
  | { readonly kind: "undefined-list"; readonly numId: number }
  /** A story was changed and no part writer carries that change into the file (`unsupported-content`) */
  | {
      readonly kind: "unwritten-story-change";
      readonly story: ExportProblemStory;
      readonly change: "added" | "edited" | "removed";
    }
  /** A story was added under an id its part identifies entries by no whole number of (`unsupported-content`) */
  | {
      readonly kind: "story-id-not-a-number";
      readonly story: ExportProblemStory;
    }
  /** A part the writer has to add cannot be declared, the package having no content types part (`missing-content-types`) */
  | {
      readonly kind: "missing-content-types";
      readonly part: ExportPartName;
    };

/** One reason the document cannot be written back, with the code `exportDocx` would throw it under */
export interface ExportProblem {
  readonly code: DocxExportErrorCode;
  /** An English one-liner for a developer; reworded freely. Read `code` and `reason` instead */
  readonly message: string;
  /** Which of the situations the code covers was met, and what it is about */
  readonly reason: ExportProblemReason;
  /**
   * Where the problem stands in the document: the block or marker it is about, or, for a footnote
   * or an endnote, the first reference to that note in the body. Absent for a problem of the
   * package, of the session, of a header or footer, and of a note the body refers to nowhere.
   */
  readonly pos?: number;
}

/** What a `DocxExportError` may be given beyond its code and message */
export interface DocxExportErrorOptions extends ErrorOptions {
  /** The entry `exportProblems` reports for this refusal, where an invariant predicted it */
  readonly problem?: ExportProblem;
}

/** Thrown when an edited document cannot be written back out without losing or corrupting content */
export class DocxExportError extends Error {
  readonly code: DocxExportErrorCode;
  /**
   * What `exportProblems` reports for this refusal, carrying its `reason` and `pos`. Absent for a
   * refusal met while writing that no invariant predicted.
   */
  readonly problem?: ExportProblem;

  constructor(
    code: DocxExportErrorCode,
    message: string,
    options?: DocxExportErrorOptions
  ) {
    super(message, options);
    this.name = "DocxExportError";
    this.code = code;
    this.problem = options?.problem;
  }
}
