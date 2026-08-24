/**
 * Reads docx bytes and builds the document to edit plus the store holding the original.
 *
 * This module knows nothing about the screen. All it knows is the file and the document model.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import { type RunFormat, toRunFormat } from "../model/format";
import { parseNumbering } from "../numbering/parseNumbering";
import { DocxImportError } from "../ooxml/errors";
import {
  childByLocalName,
  decodeUtf8,
  elementChildren,
  localPart,
  parseXml,
  R_NS,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
import { commentReferencesIn, readComments } from "./comments";
import { openParts } from "./container";
import { DEFAULT_TAB_STOP_PT, readDefaultTabStop } from "./documentSettings";
import {
  defaultParagraphStyleIdOf,
  defaultTableStyleIdOf,
  effectiveParagraphFormat,
  effectiveParagraphStyle,
  layerRunFormat,
  NO_DOCUMENT_DEFAULTS,
  NO_STYLES,
  type ParagraphFormattingContext,
  readDefaultParagraphFormat,
  readDocumentDefaults,
  readParagraphStyles,
  readStyles,
  type StyleTable,
} from "./formatting";
import { readHeadersFooters } from "./headersFooters";
import { readLinkTargets } from "./hyperlink";
import { buildParagraph, type ImportSources } from "./importParagraph";
import { buildTable } from "./importTable";
import { readImageSources } from "./media";
import { readNotes } from "./notes";
import { readBodyGeometry } from "./pageGeometry";
import { resolveTarget } from "./relationships";
import { type BodyScan, scanBody } from "./scan";
import { SessionStore } from "./session";
import { NO_THEME_FONTS, readThemeFonts } from "./theme";

const OFFICE_DOCUMENT_REL = `${R_NS}/officeDocument`;
const STYLES_REL = `${R_NS}/styles`;
const NUMBERING_REL = `${R_NS}/numbering`;
const THEME_REL = `${R_NS}/theme`;
const SETTINGS_REL = `${R_NS}/settings`;

/** Finds the target path of the given relationship type in a relationship file inside the zip */
function relationshipTarget(
  parts: Map<string, Uint8Array>,
  relsPath: string,
  type: string
): string | null {
  const rels = parts.get(relsPath);
  if (!rels) return null;
  for (const rel of elementChildren(
    parseXml(decodeUtf8(rels).text).documentElement
  )) {
    if (rel.getAttribute("Type") === type) return rel.getAttribute("Target");
  }
  return null;
}

/** Finds where inside the zip the part holding the body sits */
function findMainPartPath(parts: Map<string, Uint8Array>): string {
  const target = relationshipTarget(parts, "_rels/.rels", OFFICE_DOCUMENT_REL);
  if (!target) {
    throw new DocxImportError(
      "missing-part",
      "no relationship pointing at the main document part"
    );
  }
  return target.replace(/^\//, "");
}

/** Where inside the zip a part paired with the body part sits. null if there is no such relationship */
function relatedPartPath(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  type: string
): string | null {
  const directory = mainPartPath.replace(/[^/]+$/, "");
  const fileName = mainPartPath.slice(directory.length);
  const target = relationshipTarget(
    parts,
    `${directory}_rels/${fileName}.rels`,
    type
  );
  if (!target) return null;
  return resolveTarget(mainPartPath, target);
}

/** Pulls a single part out as text. This is the only place that knows about the zip */
function readPart(
  parts: Map<string, Uint8Array>,
  path: string | null
): string | null {
  const bytes = path === null ? undefined : parts.get(path);
  return bytes ? decodeUtf8(bytes).text : null;
}

/** Moves a single body block into a node. If we cannot model it, the result is a preservation node pointing at the original fragment */
function buildBlock(
  el: Element,
  srcId: number,
  sources: ImportSources,
  styles: StyleTable,
  defaultTableStyleId: string | null
): PMNode {
  if (el.localName === "bookmarkStart" || el.localName === "bookmarkEnd") {
    return docxSchema.nodes.bookmarkBlock.create({
      srcId,
      name: el.nodeName,
    });
  }
  if (el.localName === "p") {
    const paragraph = buildParagraph(el, srcId, sources);
    if (paragraph) return paragraph;
  }
  if (el.localName === "tbl") {
    const table = buildTable(el, srcId, sources, styles, defaultTableStyleId);
    if (table) return table;
  }
  return docxSchema.nodes.docxRaw.create({ srcId, name: el.nodeName });
}

/** Lays the style values underneath the display values of the run mark attached to a piece of text */
function styledInline(node: PMNode, style: RunFormat): PMNode {
  const mark = node.marks.find((entry) => entry.type === docxSchema.marks.run);
  if (!mark) return node;
  const format = layerRunFormat(style, toRunFormat(mark.attrs.format));
  const next = mark.type.create({ ...mark.attrs, format });
  return node.mark(next.addToSet(node.marks));
}

/**
 * Lays the values of the style a paragraph wears underneath the display values of the paragraph
 * and of the text inside it.
 *
 * The style's run values are also baked onto the paragraph itself, so that text carrying no run
 * of its own - typed in the editor - is drawn in them (`styleRun` in `schema`).
 */
function styledParagraph(
  node: PMNode,
  context: ParagraphFormattingContext
): PMNode {
  const style = effectiveParagraphStyle(node.attrs.pPr, context);
  const inline = style
    ? node.children.map((child) => styledInline(child, style.run))
    : node.children;
  return node.type.create(
    {
      ...node.attrs,
      format: effectiveParagraphFormat(node.attrs.pPr, context),
      styleRun: style ? layerRunFormat(style.run, null) : null,
    },
    Fragment.fromArray(inline),
    node.marks
  );
}

/**
 * Folds the style chain into the display values.
 *
 * These values are used for display only, so the original XML fragments are left untouched.
 * We walk down through the blocks so that paragraphs inside table cells take the same path.
 */
function withStyleFormats(
  node: PMNode,
  context: ParagraphFormattingContext
): PMNode {
  if (node.type === docxSchema.nodes.paragraph) {
    return styledParagraph(node, context);
  }
  if (node.childCount === 0) return node;
  const children = node.children.map((child) =>
    withStyleFormats(child, context)
  );
  // If no child changed, do not rebuild the node
  if (children.every((child, i) => child === node.child(i))) return node;
  return node.copy(Fragment.fromArray(children));
}

/**
 * Checks that the fragments sliced out of the raw text and what the DOM read point at the same thing.
 * If the two disagree, even the parts we never edited could be corrupted, so we do not open the file.
 */
function assertScanMatchesDom(children: Element[], scan: BodyScan): void {
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
 * Folds the `w:sectPr` at the very end of the body into the tail instead of making it a document node.
 *
 * The page setup is not visible in Word's body either, and there is nothing about it to edit.
 * Left as a document node it would disappear on select-all + delete, and whatever was written
 * afterwards would go out with no page setup at all.
 * Attached to the tail, the byte sequence stays as it was, so an edit-free round trip is undisturbed too.
 *
 * A sectPr anywhere other than the very end is a shape no healthy document has, so it is simply
 * left as a preservation block.
 */
function foldTrailingSectPr(scan: BodyScan): BodyScan {
  const last = scan.blocks.at(-1);
  // In a document whose body is nothing but a sectPr, folding it away would leave no block to edit
  if (scan.blocks.length < 2 || !last) return scan;
  if (localPart(last.name) !== "sectPr") return scan;
  return {
    prefix: scan.prefix,
    blocks: scan.blocks.slice(0, -1),
    suffix: last.xml + scan.suffix,
  };
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

export function importDocx(input: DocxBytes): {
  doc: PMNode;
  session: SessionStore;
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
  const body = childByLocalName(dom.documentElement, "body");
  if (!body) {
    throw new DocxImportError("missing-body", "document has no w:body");
  }

  const children = elementChildren(body);
  assertScanMatchesDom(children, scanned);

  const geometry = readBodyGeometry(body);
  const scan = foldTrailingSectPr(scanned);
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
  const defaultTabStopPt =
    readDefaultTabStop(settingsXml === null ? null : parseXml(settingsXml)) ??
    DEFAULT_TAB_STOP_PT;
  const numberingPartPath = relatedPartPath(parts, mainPartPath, NUMBERING_REL);
  const themeXml = readPart(
    parts,
    relatedPartPath(parts, mainPartPath, THEME_REL)
  );
  // Display only: a run pointing at a theme font is drawn in it, and carries the
  // reference itself back out untouched
  const themeFonts =
    themeXml === null ? NO_THEME_FONTS : readThemeFonts(parseXml(themeXml));
  const styles = stylesDom ? readStyles(stylesDom, themeFonts) : NO_STYLES;
  const defaultTableStyleId = stylesDom
    ? defaultTableStyleIdOf(stylesDom)
    : null;
  const defaultParagraphStyleId = stylesDom
    ? defaultParagraphStyleIdOf(stylesDom)
    : null;
  const paragraphDefaults = stylesDom
    ? readDefaultParagraphFormat(stylesDom)
    : {};
  const numberingXml = readPart(parts, numberingPartPath);
  const numbering = parseNumbering(numberingXml);
  const paragraphFormatting: ParagraphFormattingContext = {
    styles,
    defaultStyleId: defaultParagraphStyleId,
    defaults: paragraphDefaults,
    numbering,
  };
  const comments = readComments(parts, mainPartPath);
  const notes = readNotes(parts, mainPartPath);
  const noteLabels = {
    footnote: new Map<string, string>(),
    endnote: new Map<string, string>(),
  };
  const noteLabel = (kind: "footnote" | "endnote", id: string): string => {
    const labels = noteLabels[kind];
    const existing = labels.get(id);
    if (existing) return existing;
    const label = `${labels.size + 1}`;
    labels.set(id, label);
    return label;
  };
  const headersFooters = readHeadersFooters(parts, mainPartPath, body);
  const sources: ImportSources = {
    images: readImageSources(parts, mainPartPath),
    themeFonts,
    links: readLinkTargets(parts, mainPartPath),
    comments,
    notes,
    noteLabel,
  };
  const blockNodes = blockElements.map((el, i) =>
    withStyleFormats(
      buildBlock(el, i, sources, styles, defaultTableStyleId),
      paragraphFormatting
    )
  );
  return {
    doc: docxSchema.nodes.doc.create(null, blockNodes),
    session: new SessionStore({
      parts,
      mainPartPath,
      documentPrefix: scan.prefix,
      documentSuffix: scan.suffix,
      documentHadBom: hadBom,
      blocks: blockNodes.map((node, i) => ({ xml: scan.blocks[i].xml, node })),
      defaults: stylesDom
        ? readDocumentDefaults(stylesDom, themeFonts)
        : NO_DOCUMENT_DEFAULTS,
      defaultTabStopPt,
      paragraphDefaults,
      geometry,
      styles,
      paragraphStyles: stylesDom ? readParagraphStyles(stylesDom) : [],
      defaultParagraphStyleId,
      numberingXml,
      numberingPartPath,
      comments,
      commentReferenceIds: new Set(
        commentReferencesIn(
          docxSchema.nodes.doc.create(null, blockNodes)
        ).keys()
      ),
      headersFooters,
    }),
  };
}
