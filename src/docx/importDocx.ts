/**
 * Reads docx bytes and builds the document to edit plus the store holding the original.
 *
 * This module knows nothing about the screen. All it knows is the file and the document model.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  bindsWritingPrefix,
  conformanceOf,
  STRICT_OFFICE_DOCUMENT_REL,
} from "../ooxml/conformance";
import { DocxImportError } from "../ooxml/errors";
import {
  childByLocalName,
  decodeUtf8,
  elementChildren,
  localPart,
  parseXml,
  R_NS,
  withXmlParser,
  type XmlParser,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  commentReferencesIn,
  type ImportedComments,
  readComments,
} from "./comments";
import { openParts } from "./container";
import {
  DEFAULT_TAB_STOP_PT,
  readCompatSettings,
  readDefaultTabStop,
} from "./documentSettings";
import { type FidelityNote, fidelityNotesOf } from "./fidelity";
import {
  type FormattingContext,
  formattingContextOf,
  NO_DOCUMENT_DEFAULTS,
  readDocumentDefaults,
  readParagraphStyles,
} from "./formatting";
import { readHeaderFooterStories } from "./headersFooters";
import { readLinkTargets } from "./hyperlink";
import { type ImportSources, NO_IMPORT_SOURCES } from "./importParagraph";
import { readImageSources } from "./media";
import { NUMBERING_REL_TYPE } from "./newLists";
import { withNoteLabels } from "./notes/numbering";
import {
  type ImportedNotePart,
  type NoteKind,
  readNoteNumbering,
  readNotes,
  specialNotesOf,
} from "./notes/reading";
import { readPart, relatedPartPath } from "./packageParts";
import { A4_PORTRAIT } from "./pageGeometry";
import { readRelationships } from "./relationships";
import { type BlockScan, scanBody } from "./scan";
import {
  firstSectPrElement,
  readSectionProperties,
  storyReferenceIds,
} from "./sections";
import {
  BODY_STORY_KEY,
  blockKey,
  newSessionId,
  SessionStore,
} from "./session";
import {
  buildBlock,
  type ImportedStory,
  readStories,
  type StoryDeps,
  storiesByKey,
  withStyleFormats,
} from "./story";
import { NO_THEME_FONTS, readThemeFonts } from "./theme";

const OFFICE_DOCUMENT_REL = `${R_NS}/officeDocument`;
const STYLES_REL = `${R_NS}/styles`;
const THEME_REL = `${R_NS}/theme`;
const SETTINGS_REL = `${R_NS}/settings`;

/** Finds where inside the zip the part holding the body sits */
function findMainPartPath(parts: Map<string, Uint8Array>): string {
  const relationships = readRelationships(parts, "_rels/.rels");
  const target = relationships.find(
    (rel) => rel.type === OFFICE_DOCUMENT_REL
  )?.target;
  if (!target) {
    // A Strict package names the same relationship under its own type, so it has a main part and
    // is refused for what it is rather than for lacking one
    if (relationships.some((rel) => rel.type === STRICT_OFFICE_DOCUMENT_REL)) {
      throw new DocxImportError(
        "unsupported-conformance",
        "the package relates its main document part as an ECMA-376 Strict one"
      );
    }
    throw new DocxImportError(
      "missing-part",
      "no relationship pointing at the main document part"
    );
  }
  return target.replace(/^\//, "");
}

/**
 * Turns down a main part this editor could read but never write back into: one written in the
 * Strict vocabulary, or one whose root does not bind `w` to the Transitional namespace.
 * Strict is asked first so that its namespace is refused under the conformance code.
 */
function assertWritableMainPart(root: Element): void {
  if (conformanceOf(root) === "strict") {
    throw new DocxImportError(
      "unsupported-conformance",
      "the main document part is written in the ECMA-376 Strict vocabulary"
    );
  }
  if (conformanceOf(root) !== "transitional" || root.localName !== "document") {
    throw new DocxImportError(
      "unsupported-content",
      "the main part is not a Transitional WordprocessingML document"
    );
  }
  if (!bindsWritingPrefix(root)) {
    throw new DocxImportError(
      "unsupported-content",
      "the main document part binds WordprocessingML to a prefix other than w"
    );
  }
}

/**
 * Checks that the fragments sliced out of the raw text and what the DOM read point at the same thing.
 * If the two disagree, even the parts we never edited could be corrupted, so we do not open the file.
 */
function assertScanMatchesDom(children: Element[], scan: BlockScan): void {
  if (children.length !== scan.blocks.length) {
    throw new DocxImportError(
      "malformed-xml",
      "the body block count differs between the scan and the DOM"
    );
  }
  children.forEach((el, i) => {
    if (el.nodeName !== scan.blocks[i].name) {
      throw new DocxImportError(
        "malformed-xml",
        `body block names disagree: ${el.nodeName} vs ${scan.blocks[i].name}`
      );
    }
  });
}

/**
 * Takes the `w:sectPr` at the very end of the body out of the blocks, so that it can be carried by
 * the document node rather than stand among the blocks an edit works on (§17.6.18).
 *
 * The page setup is not a block in Word's body either, and a block is what select-all + delete
 * takes away; on the document node it survives that and still rides the transaction that changes
 * it. The slice keeps whatever stood in front of it, so an untouched document goes back out as the
 * bytes it arrived as.
 *
 * A `w:sectPr` anywhere other than the very end is a shape no healthy document has, so it is
 * simply left as a preservation block.
 */
function splitTrailingSectPr(scan: BlockScan): BlockScan & {
  sectPr: string | null;
} {
  const last = scan.blocks.at(-1);
  if (!last || localPart(last.name) !== "sectPr") {
    return { ...scan, sectPr: null };
  }
  return {
    prefix: scan.prefix,
    blocks: scan.blocks.slice(0, -1),
    suffix: scan.suffix,
    sectPr: last.xml,
  };
}

/**
 * What reading a side story of this package takes: the images and links of the part it stands in,
 * and no comment or note of its own, since WordprocessingML puts neither inside one.
 */
function storyDeps(
  parts: Map<string, Uint8Array>,
  partPath: string,
  sessionId: string,
  formatting: FormattingContext
): StoryDeps {
  return {
    session: { sessionId },
    sources: {
      ...NO_IMPORT_SOURCES,
      images: readImageSources(parts, partPath),
      links: readLinkTargets(parts, partPath),
      themeFonts: formatting.themeFonts,
    },
    formatting,
  };
}

/** Every comment body of the package, each read the way a body block is */
function readCommentStories(
  comments: ImportedComments,
  parts: Map<string, Uint8Array>,
  sessionId: string,
  formatting: FormattingContext
): readonly ImportedStory[] {
  const { partPath, xml } = comments;
  if (partPath === null || xml === null) return [];
  return readStories(
    { kind: "comment", partPath, xml, entryName: "comment", idAttr: "id" },
    storyDeps(parts, partPath, sessionId, formatting)
  );
}

/** The same for one notes part, whose entries are its footnotes or its endnotes */
function readNoteStories(
  part: ImportedNotePart,
  kind: NoteKind,
  parts: Map<string, Uint8Array>,
  sessionId: string,
  formatting: FormattingContext
): readonly ImportedStory[] {
  const { partPath, xml } = part;
  if (partPath === null || xml === null) return [];
  return readStories(
    { kind, partPath, xml, entryName: kind, idAttr: "id" },
    storyDeps(parts, partPath, sessionId, formatting)
  );
}

/** A docx file as bytes, however it arrived (file input, fetch, filesystem) */
export type DocxBytes = ArrayBuffer | Uint8Array;

/**
 * A docx file however a consumer holds it: the bytes, or the `File`/`Blob` a file
 * input or a fetch hands over.
 *
 * Reading a `Blob` is asynchronous, so only the React component takes one; the
 * engine on `./core` stays synchronous and bytes-only.
 */
export type DocxSource = DocxBytes | Blob;

/** What a caller may say about a read beyond handing over the bytes */
export interface ImportOptions {
  /**
   * The parser every part of the package is read through. Left out, the `DOMParser` global is
   * used, and a runtime carrying none refuses the read with `no-xml-parser`.
   */
  xmlParser?: XmlParser;
}

export function importDocx(
  input: DocxBytes,
  options?: ImportOptions
): {
  doc: PMNode;
  session: SessionStore;
  notes: FidelityNote[];
} {
  return withXmlParser(options?.xmlParser, () => readDocx(input));
}

function readDocx(input: DocxBytes): {
  doc: PMNode;
  session: SessionStore;
  notes: FidelityNote[];
} {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const parts = openParts(bytes);
  const mainPartPath = findMainPartPath(parts);
  const mainPart = parts.get(mainPartPath);
  if (!mainPart) {
    throw new DocxImportError("missing-part", `${mainPartPath} is missing`);
  }

  const { text: source, hadBom } = decodeUtf8(mainPart);
  const scanned = scanBody(source);
  const dom = parseXml(source);
  assertWritableMainPart(dom.documentElement);
  const body = childByLocalName(dom.documentElement, "body");
  if (!body) {
    throw new DocxImportError("missing-body", "document has no w:body");
  }

  const children = elementChildren(body);
  assertScanMatchesDom(children, scanned);

  const firstSectPr = firstSectPrElement(body);
  const firstSection = firstSectPr ? readSectionProperties(firstSectPr) : null;
  const geometry = firstSection?.geometry ?? A4_PORTRAIT;
  const scan = splitTrailingSectPr(scanned);
  const blockElements = children.slice(0, scan.blocks.length);

  const stylesXml = readPart(
    parts,
    relatedPartPath(parts, mainPartPath, STYLES_REL)
  );
  const stylesDom = stylesXml === null ? null : parseXml(stylesXml);
  const settingsXml = readPart(
    parts,
    relatedPartPath(parts, mainPartPath, SETTINGS_REL)
  );
  const settingsDom = settingsXml === null ? null : parseXml(settingsXml);
  const defaultTabStopPt =
    readDefaultTabStop(settingsDom) ?? DEFAULT_TAB_STOP_PT;
  const numberingPartPath = relatedPartPath(
    parts,
    mainPartPath,
    NUMBERING_REL_TYPE
  );
  const themeXml = readPart(
    parts,
    relatedPartPath(parts, mainPartPath, THEME_REL)
  );
  // Display only: a run pointing at a theme font is drawn in it, and carries the
  // reference itself back out untouched
  const themeFonts =
    themeXml === null ? NO_THEME_FONTS : readThemeFonts(parseXml(themeXml));
  const numberingXml = readPart(parts, numberingPartPath);
  const formatting = formattingContextOf(
    stylesDom,
    numberingXml,
    themeFonts,
    readCompatSettings(settingsDom)
  );
  const comments = readComments(parts, mainPartPath);
  const notes = readNotes(parts, mainPartPath);
  const noteNumbering = readNoteNumbering(settingsDom);
  const specialNotes = specialNotesOf(notes);
  const sources: ImportSources = {
    images: readImageSources(parts, mainPartPath),
    themeFonts,
    links: readLinkTargets(parts, mainPartPath),
    comments,
  };
  const sessionId = newSessionId();
  const headerFooters = readHeaderFooterStories(
    parts,
    mainPartPath,
    storyReferenceIds(body),
    (partPath) => storyDeps(parts, partPath, sessionId, formatting)
  );
  const stories = storiesByKey([
    ...readCommentStories(comments, parts, sessionId, formatting),
    ...readNoteStories(
      notes.footnotes,
      "footnote",
      parts,
      sessionId,
      formatting
    ),
    ...readNoteStories(notes.endnotes, "endnote", parts, sessionId, formatting),
    ...headerFooters.stories,
  ]);
  const blockNodes = blockElements.map((el, i) =>
    withStyleFormats(
      buildBlock(
        el,
        blockKey({ sessionId }, BODY_STORY_KEY, i),
        sources,
        formatting
      ),
      formatting
    )
  );
  if (blockNodes.length === 0 && scan.sectPr !== null) {
    // A section-only body needs a place to type. Its original block is empty XML, so the
    // paragraph is written only after an edit and an untouched file remains byte-identical.
    blockNodes.push(
      withStyleFormats(
        docxSchema.nodes.paragraph.create({
          srcId: blockKey({ sessionId }, BODY_STORY_KEY, 0),
        }),
        formatting
      )
    );
  }
  const doc = withNoteLabels(
    docxSchema.nodes.doc.create(
      {
        sectPr: scan.sectPr,
        stories: Object.fromEntries(
          Array.from(stories, ([key, story]) => [key, story.doc.toJSON()])
        ),
      },
      blockNodes
    ),
    noteNumbering,
    specialNotes
  );
  return {
    doc,
    notes: fidelityNotesOf(doc, mainPartPath),
    session: new SessionStore({
      sessionId,
      parts,
      mainPartPath,
      documentPrefix: scan.prefix,
      documentSuffix: scan.suffix,
      documentHadBom: hadBom,
      blocks: doc.children.map((node, i) => ({
        xml: scan.blocks[i]?.xml ?? "",
        node,
      })),
      defaults: stylesDom
        ? readDocumentDefaults(stylesDom, themeFonts)
        : NO_DOCUMENT_DEFAULTS,
      defaultTabStopPt,
      geometry,
      formatting,
      paragraphStyles: stylesDom ? readParagraphStyles(stylesDom) : [],
      numberingXml,
      numberingPartPath,
      comments,
      commentReferenceIds: new Set(commentReferencesIn(doc).keys()),
      headerFooterStories: headerFooters.refs,
      stories,
      noteNumbering,
      specialNotes,
    }),
  };
}
