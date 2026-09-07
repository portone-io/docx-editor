/**
 * The store that holds the original as it was the moment the document was opened
 * and hands it back at export time.
 *
 * What leaves the package is `DocxSession`, a token with nothing readable on it: the store behind
 * it holds the file's own XML, which nothing but the exporter can put back together, and a
 * consumer reading a piece of it would be reading an internal shape. What a consumer legitimately
 * needs is an accessor of its own here.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { DocumentDefaults } from "../model/format";
import {
  type Numbering,
  type NumberingOptions,
  parseNumbering,
} from "../numbering/parseNumbering";
import type { ImportedComments } from "./comments";
import type {
  ParagraphFormatLayer,
  ParagraphStyleOption,
  StyleTable,
} from "./formatting";
import type { HeadersFooters } from "./headersFooters";
import type { PageGeometry } from "./pageGeometry";

export interface ImportedBlock {
  xml: string;
  /** The node as it was the moment the document was opened. Compared against the current node to tell whether it was edited */
  node: PMNode;
}

/**
 * The session `importDocx` hands out and `exportDocx` takes back, untouched.
 *
 * It carries the parts of the package and the original XML pieces that go back out unchanged,
 * which is how a document nobody edited exports part for part identical.
 */
export interface DocxSession {
  readonly kind: "docxSession";
}

/** What names one open document, which a block key is written against */
export interface SessionIdentity {
  readonly sessionId: string;
}

let sessionsOpened = 0;

/**
 * A name for one open document.
 *
 * Opening the same bytes twice opens two documents, so this is drawn fresh each time rather than
 * worked out from what was opened. It carries no `:`, which is what lets a block key be read back
 * however many colons the story it names holds.
 */
export function newSessionId(): string {
  sessionsOpened += 1;
  // The counter alone would repeat across two copies of this module, a worker's among them
  return `d${sessionsOpened}-${Math.random().toString(36).slice(2, 10)}`;
}

/** What the package's own layers read the open document off */
export class SessionStore implements DocxSession, SessionIdentity {
  readonly kind = "docxSession";
  /** Which open document this is. Every block key of it names it, so a block of another document is not taken for one of this one */
  readonly sessionId: string;
  readonly parts: Map<string, Uint8Array>;
  readonly mainPartPath: string;
  readonly documentPrefix: string;
  readonly documentSuffix: string;
  readonly documentHadBom: boolean;
  readonly blocks: ImportedBlock[];
  /** The document default formatting read from styles.xml. Used for display only */
  readonly defaults: DocumentDefaults;
  /** The effective automatic tab interval read from settings.xml. Used for display only. */
  readonly defaultTabStopPt: number;
  /** The paragraph properties at the base of the OOXML formatting hierarchy. */
  readonly paragraphDefaults: ParagraphFormatLayer;
  /** The paper this document is written on, read from the first section. Used for display only: the `w:sectPr` itself goes back out in the preserved tail */
  readonly geometry: PageGeometry;
  /** The style chain from styles.xml. Used only to fold the style a paragraph points at into its displayed values */
  readonly styles: StyleTable;
  /** The paragraph styles this document defines, in the order styles.xml lists them. What the style picker offers */
  readonly paragraphStyles: ParagraphStyleOption[];
  /**
   * The paragraph style every paragraph with no `w:pStyle` of its own wears (`w:default="1"`,
   * usually Normal). null when the document marks none
   */
  readonly defaultParagraphStyleId: string | null;
  /** The raw text of numbering.xml. Kept around so list numbers can be drawn on screen. null if there is none */
  readonly numberingXml: string | null;
  /** Where numbering.xml sits inside the zip. When a new list is exported it is rewritten at that spot. null if there is none */
  readonly numberingPartPath: string | null;
  /** The Comments part as opened, including comment elements retained for a byte-identical pass. */
  readonly comments: ImportedComments;
  /** Comment ids referenced by the original main story. A missing id after editing means deletion. */
  readonly commentReferenceIds: ReadonlySet<string>;
  /** First-section header and footer stories projected for the page preview. */
  readonly headersFooters: HeadersFooters;

  constructor(opened: Omit<SessionStore, "kind">) {
    this.sessionId = opened.sessionId;
    this.parts = opened.parts;
    this.mainPartPath = opened.mainPartPath;
    this.documentPrefix = opened.documentPrefix;
    this.documentSuffix = opened.documentSuffix;
    this.documentHadBom = opened.documentHadBom;
    this.blocks = opened.blocks;
    this.defaults = opened.defaults;
    this.defaultTabStopPt = opened.defaultTabStopPt;
    this.paragraphDefaults = opened.paragraphDefaults;
    this.geometry = opened.geometry;
    this.styles = opened.styles;
    this.paragraphStyles = opened.paragraphStyles;
    this.defaultParagraphStyleId = opened.defaultParagraphStyleId;
    this.numberingXml = opened.numberingXml;
    this.numberingPartPath = opened.numberingPartPath;
    this.comments = opened.comments;
    this.commentReferenceIds = opened.commentReferenceIds;
    this.headersFooters = opened.headersFooters;
  }
}

/**
 * The store behind a session token.
 *
 * Every function that takes a session from outside comes through here, so an object the package
 * never handed out is turned down with a message rather than read field by field into nonsense.
 */
export function sessionOf(session: DocxSession): SessionStore {
  if (session instanceof SessionStore) return session;
  throw new Error("the session must be the one importDocx handed back");
}

/** Where the body part sits inside the package, which is the one part an export rewrites */
export function documentPartPath(session: DocxSession): string {
  return sessionOf(session).mainPartPath;
}

/** The list definitions the document carries, which is what a paragraph's numbering points into */
export function documentNumbering(
  session: DocxSession,
  options?: NumberingOptions
): Numbering {
  return parseNumbering(sessionOf(session).numberingXml, options);
}

/** The story a block key names when the block stood in the main document body */
export const BODY_STORY_KEY = "body";

/** A block key read back: which open document the block came from, which story of it, and where in that story */
export interface BlockKey {
  sessionId: string;
  storyKey: string;
  index: number;
}

/** Where a block stands, as the one string a node carries on its `srcId` */
export function blockKey(
  session: SessionIdentity,
  storyKey: string,
  index: number
): string {
  return `${session.sessionId}:${storyKey}:${index}`;
}

/**
 * The three parts of a block key, or null for a string that is not one.
 *
 * A story key holds colons of its own (`comment:4`), so the key is read from both ends rather
 * than split: the first segment is the document, the last is the place in the story, and
 * everything between them is the story itself.
 */
export function splitBlockKey(key: string): BlockKey | null {
  const first = key.indexOf(":");
  const last = key.lastIndexOf(":");
  if (first < 1 || last <= first) return null;
  const index = key.slice(last + 1);
  if (!/^\d+$/.test(index)) return null;
  return {
    sessionId: key.slice(0, first),
    storyKey: key.slice(first + 1, last),
    index: Number(index),
  };
}

/**
 * Finds which original block a node came from. undefined for a node newly created during editing,
 * and for one that came in from another document, whose original is in that file and not this one.
 */
export function originalBlock(
  node: PMNode,
  session: SessionStore
): ImportedBlock | undefined {
  const srcId: unknown = node.attrs.srcId;
  if (typeof srcId !== "string") return undefined;
  const key = splitBlockKey(srcId);
  if (key === null || key.sessionId !== session.sessionId) return undefined;
  // The body is the only story a session holds blocks of
  return key.storyKey === BODY_STORY_KEY
    ? session.blocks[key.index]
    : undefined;
}
