// @vitest-environment jsdom
/**
 * The guard that what the export writes is WordprocessingML the standard recognises.
 *
 * Every fixture goes out untouched, with a paragraph rewritten, and once more after the edit
 * battery below, and every WordprocessingML part of the package that comes back is validated
 * against the transitional schemas committed under `spec/schemas/`. The validator is xmllint,
 * which macOS ships and a Linux image installs as libxml2-utils; a missing one fails the run
 * rather than skipping it.
 *
 * The environment is jsdom rather than node because reading a docx needs a DOMParser.
 */

import { type SpawnSyncReturns, spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";
import { afterAll, describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  decodeBase64,
  fixtureNames,
  makeDocx,
  makeHeadersFootersDocx,
  makeNotesDocx,
  readFixture,
  TINY_PNG,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import {
  type ActiveParagraphAlign,
  activeLineSpacing,
  activeParagraphAlign,
  addComment,
  type ImageToInsert,
  insertImage,
  insertTable,
  lockSelection,
  setLineSpacing,
  setLink,
  setParagraphAlign,
  setTextColor,
  toggleBold,
  toggleBulletList,
  toggleNumberedList,
  unlockSelection,
} from "../editor/commands";
import { createEditorState } from "../editor/createEditor";
import {
  type LineSpacing,
  type ParagraphAlign,
  toParagraphFormat,
} from "../model/format";
import { wAttr } from "../ooxml/units";
import { childByLocalName, decodeUtf8, parseXml, W_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  addRowAfter,
  mergeCells,
  setCellPadding,
  setCellVerticalAlign,
} from "../table";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { documentNumbering, type SessionStore } from "./session";

const XSD_NS = "http://www.w3.org/2001/XMLSchema";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const wmlPath = join(
  dirname(fileURLToPath(import.meta.url)),
  "../../spec/schemas/transitional/wml.xsd"
);

/**
 * Every schema that names `xml:space` imports the XML namespace with no schemaLocation, and a
 * validator that cannot resolve it refuses to compile the set at all. `xml:space` is the only
 * attribute of that namespace the transitional schemas refer to.
 */
const XML_NAMESPACE_SCHEMA =
  `<xsd:schema xmlns:xsd="${XSD_NS}" targetNamespace="${XML_NS}">` +
  '<xsd:attribute name="space"><xsd:simpleType>' +
  '<xsd:restriction base="xsd:NCName">' +
  '<xsd:enumeration value="default"/><xsd:enumeration value="preserve"/>' +
  "</xsd:restriction></xsd:simpleType></xsd:attribute></xsd:schema>";

/**
 * What xmllint is pointed at. wml.xsd stays where it lies, so that the relative imports of its
 * sibling schemas resolve; this names it alongside the XML namespace and nothing else.
 */
const ENTRY_SCHEMA =
  `<xsd:schema xmlns:xsd="${XSD_NS}">` +
  `<xsd:import namespace="${XML_NS}" schemaLocation="xml.xsd"/>` +
  `<xsd:import namespace="${W_NS}" schemaLocation="${pathToFileURL(wmlPath).href}"/>` +
  "</xsd:schema>";

const schemaDir = mkdtempSync(join(tmpdir(), "docx-editor-schema-"));
const entryPath = join(schemaDir, "entry.xsd");
writeFileSync(join(schemaDir, "xml.xsd"), XML_NAMESPACE_SCHEMA);
writeFileSync(entryPath, ENTRY_SCHEMA);

afterAll(() => rmSync(schemaDir, { recursive: true, force: true }));

interface Validation {
  valid: boolean;
  report: string;
}

/** What xmllint says of a document it read right through and found nothing to fault */
const ACCEPTED = "validates";

/**
 * The report of every part xmllint turned down, keyed by the path it named the part by.
 *
 * A part counts as accepted only where xmllint said so of it by name: a document it could not
 * parse at all draws a parser error and neither verdict, and reading silence as a pass would
 * hand this suite a way to stay green on a part nothing ever validated.
 */
function rejectionsIn(
  report: string,
  paths: Iterable<string>
): Map<string, string> {
  const lines = report.split("\n");
  const rejections = new Map<string, string>();
  for (const path of paths) {
    if (lines.includes(`${path} ${ACCEPTED}`)) continue;
    const named = lines.filter((line) => line.startsWith(`${path}`));
    rejections.set(path, named.length > 0 ? named.join("\n") : report.trim());
  }
  return rejections;
}

/**
 * Validates a whole export in one xmllint run.
 *
 * Compiling the transitional schema set is what a run costs - two orders of magnitude over
 * reading a part against it - and one invocation compiles it once however many documents it is
 * handed. The parts go to a temp directory at their own paths and are named to xmllint relative
 * to it, so every line it reports back is prefixed by the part path it belongs to.
 */
function validateParts(parts: Map<string, string>): Map<string, string> {
  const partsDir = mkdtempSync(join(tmpdir(), "docx-editor-parts-"));
  try {
    for (const [path, xml] of parts) {
      const file = join(partsDir, path);
      mkdirSync(dirname(file), { recursive: true });
      writeFileSync(file, xml);
    }
    const run = spawnSync(
      "xmllint",
      ["--noout", "--schema", entryPath, ...parts.keys()],
      { cwd: partsDir, encoding: "utf8" }
    );
    if (run.error) throw spawnFailure(run.error);
    // xmllint exits non-zero for a part it turned down and says on stderr which and why. One that
    // ends with nothing to say never validated anything, and reading that silence as a rejection
    // would let the negative control below pass without a validator having run at all
    if (run.status !== 0 && run.stderr.trim() === "") {
      throw new Error(
        `xmllint ${endOf(run)} with nothing on stderr, so no part was validated`
      );
    }
    return rejectionsIn(run.stderr, parts.keys());
  } finally {
    rmSync(partsDir, { recursive: true, force: true });
  }
}

function errorCode(error: Error): string | undefined {
  return "code" in error && typeof error.code === "string"
    ? error.code
    : undefined;
}

/** A validator that could not be started at all: missing is the one case worth naming */
function spawnFailure(error: Error): Error {
  if (errorCode(error) === "ENOENT") {
    return new Error(
      `this test validates against the OOXML schemas with xmllint, which is not on PATH: ${error.message}`
    );
  }
  return new Error(`xmllint could not be run: ${error.message}`);
}

function endOf(run: SpawnSyncReturns<string>): string {
  return run.signal === null
    ? `exited with ${run.status}`
    : `was killed by ${run.signal}`;
}

/** One document read under the part path it would go out as */
function validate(path: string, xml: string): Validation {
  const report = validateParts(new Map([[path, xml]])).get(path);
  return { valid: report === undefined, report: report ?? "" };
}

/**
 * Markup compatibility (ECMA-376 part 3) has a consumer drop the compatibility attributes
 * before reading a part against the part 1 schemas, which know nothing of that namespace.
 * Word writes `mc:Ignorable` on the root of every part it saves, so our fixtures carry it too.
 *
 * Only attributes in that one namespace go; markup an ignorable namespace holds, such as a
 * `mc:AlternateContent` a document brought in from elsewhere, is left standing and fails
 * validation, which is the report we want rather than a quiet pass.
 */
function withoutCompatibilityAttributes(xml: string): string {
  const prefix = new RegExp(`xmlns:([\\w-]+)="${MC_NS}"`).exec(xml)?.[1];
  if (prefix === undefined) return xml;
  return xml.replaceAll(new RegExp(`\\s${prefix}:[A-Za-z]+="[^"]*"`, "g"), "");
}

/** The parts of an exported package that WordprocessingML describes, ready to validate */
function wordprocessingParts(bytes: Uint8Array): Map<string, string> {
  const parts = new Map<string, string>();
  for (const [path, data] of Object.entries(unzipSync(bytes))) {
    if (!path.endsWith(".xml")) continue;
    const xml = decode(data);
    if (parseXml(xml).documentElement.namespaceURI === W_NS) {
      parts.set(path, withoutCompatibilityAttributes(xml));
    }
  }
  return parts;
}

function expectPartsValidate(name: string, parts: Map<string, string>): void {
  expect(parts.size).toBeGreaterThan(0);
  const rejected = Array.from(
    validateParts(parts),
    ([path, report]) => `${name} ${path}\n${report}`
  );
  expect(rejected, `${name}: parts the schemas turned down`).toEqual([]);
}

const EDITED = "edited before the export was validated";

function withEditedText(paragraph: PMNode): PMNode {
  const inline: PMNode[] = [];
  let edited = false;
  paragraph.forEach((child) => {
    if (!edited && child.isText) {
      inline.push(docxSchema.text(EDITED, child.marks));
      edited = true;
    } else {
      inline.push(child);
    }
  });
  return paragraph.copy(Fragment.from(inline));
}

/**
 * The document with the first paragraph holding text rewritten, so that the part the validator
 * reads is one the export built rather than one it handed back untouched.
 */
function withEditedParagraph(doc: PMNode): PMNode {
  let target = -1;
  doc.forEach((block, _offset, index) => {
    const editable =
      block.type.name === "paragraph" && block.textContent !== "";
    if (target === -1 && editable) target = index;
  });
  if (target === -1) throw new Error("the fixture has no paragraph to edit");

  const blocks: PMNode[] = [];
  doc.forEach((block, _offset, index) => {
    blocks.push(index === target ? withEditedText(block) : block);
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

/**
 * The battery of edits that runs before the third export.
 *
 * Rewriting a paragraph reaches the paragraph serializer and nothing else: everything the export
 * writes only for content that was not in the file already - a list definition spliced into
 * numbering.xml, a table built from the template, a media part with its relationship and content
 * type, a hyperlink with the external relationship its address lives on, a content control around a
 * locked stretch - stays out of reach. These steps reach them.
 *
 * Each step runs one command of the public surface (`./commands`, `./table`) over an
 * `EditorState` with no view, and refuses to go on when the command reports it changed nothing.
 * A battery that stopped reaching a writer therefore fails here rather than quietly validating
 * an export that never ran it.
 */

const TEXT_COLOR = "#1F4E79";

/** The address the battery links a stretch of text to, which the export writes a relationship for */
const LINK_ADDRESS = "https://example.com/battery?a=1&b=2";

const CONTENT_TYPES_PATH = "[Content_Types].xml";

const A_PICTURE: ImageToInsert = {
  src: TINY_PNG_DATA_URL,
  extent: { cx: 952500, cy: 952500 },
  alt: "the picture the battery inserted",
};

/** A 1x1 transparent GIF, an image of a kind no fixture's content types declare */
const TINY_GIF_BASE64 =
  "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";

const TINY_GIF = decodeBase64(TINY_GIF_BASE64);

/**
 * A second picture of that other kind, so the export has to add the declaration rather than find
 * it there already, which is the only way the content types writer runs at all
 */
const ANOTHER_PICTURE: ImageToInsert = {
  src: `data:image/gif;base64,${TINY_GIF_BASE64}`,
  extent: { cx: 476250, cy: 476250 },
  alt: "the second picture the battery inserted",
};

interface Spot {
  pos: number;
  node: PMNode;
}

/** The editing state a screen would hold for this document */
function openState(doc: PMNode, session: SessionStore): EditorState {
  return createEditorState(doc, {
    numbering: documentNumbering(session),
    styles: session.styles,
    defaults: session.defaults,
    canStartNewList: session.numberingPartPath !== null,
    paragraphStyles: session.paragraphStyles,
  });
}

/** Runs one command over the selection the state holds, and refuses to go on when it did nothing */
function ran(state: EditorState, what: string, command: Command): EditorState {
  let next = state;
  const handled = command(state, (tr) => {
    next = state.apply(tr);
  });
  if (!handled || next === state) {
    throw new Error(`the battery could not ${what}`);
  }
  return next;
}

function bodyParagraphs(doc: PMNode): Spot[] {
  const spots: Spot[] = [];
  doc.forEach((block, offset) => {
    if (block.type.name === "paragraph")
      spots.push({ pos: offset, node: block });
  });
  return spots;
}

/**
 * Which paragraph a step works on is its place among the body's paragraphs rather than its
 * position, because the steps ahead of it insert blocks and change the formatting of others.
 * That place is the one thing none of them moves.
 */
function paragraphAt(doc: PMNode, index: number): Spot {
  const spot = bodyParagraphs(doc)[index];
  if (spot === undefined) {
    throw new Error(`the document has no paragraph ${index}`);
  }
  return spot;
}

/** The paragraphs holding text and no list marker, which is where a list toggle starts a new list */
function plainParagraphIndices(doc: PMNode): number[] {
  return bodyParagraphs(doc).flatMap(({ node }, index) => {
    const plain =
      node.textContent !== "" &&
      toParagraphFormat(node.attrs.format)?.numbering === undefined;
    return plain ? [index] : [];
  });
}

function nth(indices: readonly number[], at: number): number {
  const index = indices[at];
  if (index === undefined) {
    throw new Error(
      `the battery needs ${at + 1} plain paragraphs and the document has ${indices.length}`
    );
  }
  return index;
}

/** The caret at the start of that paragraph's text */
function caretIn(state: EditorState, index: number): EditorState {
  const { pos } = paragraphAt(state.doc, index);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, pos + 1))
  );
}

/** The whole of that paragraph's text selected */
function textOf(state: EditorState, index: number): EditorState {
  const { pos, node } = paragraphAt(state.doc, index);
  return state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, pos + 1, pos + 1 + node.content.size)
    )
  );
}

/** The table the selection sits in */
function tableAround(state: EditorState): Spot {
  const $from = state.selection.$from;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.spec.tableRole === "table") {
      return { pos: $from.before(depth), node };
    }
  }
  throw new Error("the battery left the table it inserted");
}

/** Two neighbouring cells of one row of that table selected, which is what a merge takes */
function twoCellsOfRow(state: EditorState, row: number): EditorState {
  const table = tableAround(state);
  const map = TableMap.get(table.node);
  const start = table.pos + 1;
  return state.apply(
    state.tr.setSelection(
      CellSelection.create(
        state.doc,
        start + map.positionAt(row, 0, table.node),
        start + map.positionAt(row, 1, table.node)
      )
    )
  );
}

const ALIGNS: readonly ParagraphAlign[] = [
  "center",
  "right",
  "justify",
  "left",
];

/**
 * An alignment the paragraph is not already drawn with, and a line spacing it does not already
 * carry. Both commands leave a paragraph that already reads that way untouched, so a step asking
 * for the value it already has would report it changed nothing.
 */
function otherAlign(active: ActiveParagraphAlign): ParagraphAlign {
  return (
    ALIGNS.find((align) => active.kind === "mixed" || align !== active.align) ??
    "center"
  );
}

function otherSpacing(active: LineSpacing | null): LineSpacing {
  const doubled = active?.rule === "auto" && active.lines === 2;
  return { rule: "auto", lines: doubled ? 1.5 : 2 };
}

/** The state the whole battery leaves behind */
function afterTheBattery(state: EditorState): EditorState {
  const plain = plainParagraphIndices(state.doc);
  const formatted = nth(plain, 0);
  const spaced = nth(plain, 1);
  const numbered = nth(plain, 2);
  const bulleted = nth(plain, 3);
  const tabled = nth(plain, 4);
  const pictured = nth(plain, 5);

  const bold = ran(textOf(state, formatted), "toggle bold", toggleBold);
  const colored = ran(bold, "color the text", setTextColor(TEXT_COLOR));

  const atSpaced = caretIn(colored, spaced);
  const aligned = ran(
    atSpaced,
    "align a paragraph",
    setParagraphAlign(otherAlign(activeParagraphAlign(atSpaced)))
  );
  const spacedOut = ran(
    aligned,
    "space a paragraph's lines out",
    setLineSpacing(otherSpacing(activeLineSpacing(aligned)))
  );

  const numberedList = ran(
    caretIn(spacedOut, numbered),
    "start a numbered list",
    toggleNumberedList
  );
  const bulletList = ran(
    caretIn(numberedList, bulleted),
    "start a bullet list",
    toggleBulletList
  );

  const withTable = ran(
    caretIn(bulletList, tabled),
    "insert a table",
    insertTable({ rows: 2, columns: 3 })
  );
  const withRow = ran(withTable, "add a row to that table", addRowAfter);
  const merged = ran(
    twoCellsOfRow(withRow, 1),
    "merge two cells of the added row",
    mergeCells
  );
  const verticallyAligned = ran(
    merged,
    "align selected cells vertically",
    setCellVerticalAlign("center")
  );
  const padded = ran(
    verticallyAligned,
    "pad selected cells",
    setCellPadding({ top: 6, right: 8, bottom: 6, left: 8 })
  );

  const withPicture = ran(
    caretIn(padded, pictured),
    "insert an image",
    insertImage(A_PICTURE)
  );
  const withPictures = ran(
    withPicture,
    "insert an image of another kind",
    insertImage(ANOTHER_PICTURE)
  );

  const withLink = ran(
    textOf(withPictures, formatted),
    "put a link on a stretch of text",
    setLink(LINK_ADDRESS)
  );
  const withComment = ran(
    textOf(withLink, spaced),
    "add a comment to a stretch of text",
    addComment({
      text: "The comment written by the export battery",
      author: "Schema test",
      initials: "ST",
      date: "2026-08-22T00:00:00Z",
    })
  );

  // The locks go last: the lock guard turns down every edit reaching into a locked stretch,
  // whichever step asked for it. The stretch locked first is the one now carrying a link, so the
  // export writes a control around a hyperlink as well
  const locked = ran(
    textOf(withComment, formatted),
    "lock a stretch of text",
    lockSelection
  );
  const lockedTwice = ran(
    textOf(locked, spaced),
    "lock a second stretch of text",
    lockSelection
  );
  return ran(
    textOf(lockedTwice, spaced),
    "unlock the second stretch",
    unlockSelection
  );
}

/** The same document with its tables dropped, so the one the battery inserts is the only one */
function withoutTables(doc: PMNode): PMNode {
  const blocks: PMNode[] = [];
  doc.forEach((block) => {
    if (block.type.name !== "table") blocks.push(block);
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

function partText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path];
  if (bytes === undefined) throw new Error(`the export wrote no ${path}`);
  return decodeUtf8(bytes).text;
}

function numbersOf(values: readonly (string | null)[]): number[] {
  return values.flatMap((value) => {
    const parsed = Number.parseInt(value ?? "", 10);
    return Number.isFinite(parsed) ? [parsed] : [];
  });
}

/** The list numbers numbering.xml defines */
function definedNumIds(xml: string): Set<number> {
  const nums = Array.from(parseXml(xml).getElementsByTagNameNS(W_NS, "num"));
  return new Set(numbersOf(nums.map((num) => wAttr(num, "numId"))));
}

/** The list numbers the body's paragraphs point at */
function referencedNumIds(xml: string): Set<number> {
  const refs = Array.from(parseXml(xml).getElementsByTagNameNS(W_NS, "numPr"));
  return new Set(
    numbersOf(
      refs.map((ref) => {
        const numId = childByLocalName(ref, "numId");
        return numId === null ? null : wAttr(numId, "val");
      })
    )
  );
}

/** Whether a relationships part of the package points at that media file */
function pointsAt(zip: Record<string, Uint8Array>, mediaPath: string): boolean {
  const target = `media/${mediaPath.slice(mediaPath.lastIndexOf("/") + 1)}`;
  return Object.entries(zip).some(
    ([path, bytes]) =>
      path.endsWith(".rels") && decodeUtf8(bytes).text.includes(target)
  );
}

function mediaPaths(zip: Record<string, Uint8Array>): string[] {
  return Object.keys(zip).filter((path) => path.includes("/media/"));
}

/** The one media part the export wrote for an inserted image, carrying its bytes and a relationship */
function expectMediaPart(
  name: string,
  zip: Record<string, Uint8Array>,
  added: readonly string[],
  extension: string,
  bytes: Uint8Array
): void {
  const written = added.filter((path) => path.endsWith(`.${extension}`));
  expect(written, `${name} .${extension} media parts`).toHaveLength(1);

  const path = written[0];
  expect(bytesEqual(zip[path], bytes), `${name} ${path}`).toBe(true);
  expect(pointsAt(zip, path), `${name} ${path} relationship`).toBe(true);
}

/** Every hyperlink relationship the package holds, wherever it holds it */
function linkRelationships(zip: Record<string, Uint8Array>): string[] {
  return Object.entries(zip).flatMap(([path, bytes]) =>
    path.endsWith(".rels")
      ? (decodeUtf8(bytes).text.match(/<Relationship[^>]*hyperlink[^>]*\/>/g) ??
        [])
      : []
  );
}

function lockedControlCount(xml: string): number {
  return xml.match(/sdtContentLocked/g)?.length ?? 0;
}

/** Runs the battery over one document and reads back everything the export had to write for it */
function expectBatteryValidates(
  name: string,
  doc: PMNode,
  session: SessionStore
): void {
  const bytes = exportDocx(
    afterTheBattery(openState(doc, session)).doc,
    session
  );
  expectPartsValidate(name, wordprocessingParts(bytes));

  const zip = unzipSync(bytes);
  const mainXml = partText(zip, session.mainPartPath);

  // Measured against the same document written out without the battery, because a fixture may
  // carry a picture and a link of its own
  const untouched = unzipSync(exportDocx(doc, session));

  const addedMedia = mediaPaths(zip).filter(
    (path) => !mediaPaths(untouched).includes(path)
  );
  expect(addedMedia, `${name} media parts`).toHaveLength(2);
  expectMediaPart(name, zip, addedMedia, "png", TINY_PNG);
  expectMediaPart(name, zip, addedMedia, "gif", TINY_GIF);

  const linkRels = linkRelationships(zip);
  expect(linkRels.length, `${name} hyperlink relationships`).toBe(
    linkRelationships(untouched).length + 1
  );
  const forTheBattery = linkRels.filter((rel) =>
    rel.includes('Target="https://example.com/battery?a=1&amp;b=2"')
  );
  expect(forTheBattery, `${name} hyperlink target`).toHaveLength(1);
  expect(forTheBattery[0], `${name} hyperlink relationship`).toContain(
    'TargetMode="External"'
  );

  const openedTypes = session.parts.get(CONTENT_TYPES_PATH);
  if (openedTypes === undefined) {
    throw new Error(`${name} carries no ${CONTENT_TYPES_PATH}`);
  }
  expect(
    decodeUtf8(openedTypes).text,
    `${name} declares the gif extension already, which leaves the content types writer unreached`
  ).not.toContain('Extension="gif"');
  expect(partText(zip, CONTENT_TYPES_PATH), `${name} content types`).toContain(
    'Extension="gif"'
  );

  const numberingPath = session.numberingPartPath;
  const numberingXml = session.numberingXml;
  if (numberingPath === null || numberingXml === null) {
    throw new Error(`${name} carries no numbering.xml to define a list in`);
  }
  const defined = definedNumIds(numberingXml);
  const added = Array.from(definedNumIds(partText(zip, numberingPath))).filter(
    (numId) => !defined.has(numId)
  );
  expect(added, `${name} new list definitions`).toHaveLength(2);

  const referenced = referencedNumIds(mainXml);
  expect(
    added.filter((numId) => !referenced.has(numId)),
    `${name} list definitions the body points at nowhere`
  ).toEqual([]);

  // The count to measure against is the same document written out without the battery, because
  // the caller may have dropped blocks that carried a control of their own
  expect(lockedControlCount(mainXml), `${name} locked controls`).toBe(
    lockedControlCount(partText(untouched, session.mainPartPath)) + 1
  );
}

describe("the exported package against the OOXML schemas", () => {
  it("has fixtures to export and a validator that compiles the schemas", () => {
    expect(fixtureNames.length).toBeGreaterThan(0);
    const { valid, report } = validate(
      "word/document.xml",
      `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`
    );
    expect(valid, report).toBe(true);
  });

  it.each(fixtureNames)("%s: every WordprocessingML part validates", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    expectPartsValidate(name, wordprocessingParts(exportDocx(doc, session)));
  });

  it("a body-level bookmark remains valid after a surrounding paragraph edit", () => {
    const opened = importDocx(
      makeDocx(
        '<w:bookmarkStart w:id="9" w:name="Appendix"/>' +
          '<w:p><w:r><w:t xml:space="preserve">First</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t xml:space="preserve">Second</w:t></w:r></w:p>' +
          '<w:bookmarkEnd w:id="9"/>'
      )
    );
    const parts = wordprocessingParts(
      exportDocx(withEditedParagraph(opened.doc), opened.session)
    );

    expect(parts.get(opened.session.mainPartPath)).toContain("bookmarkStart");
    expectPartsValidate("body-level bookmark", parts);
  });

  it("cell padding remains valid beside strict leading and trailing margins", () => {
    const opened = importDocx(
      makeDocx(
        "<w:tbl>" +
          '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
          '<w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>' +
          "<w:tr><w:tc><w:tcPr><w:tcMar>" +
          '<w:start w:w="80" w:type="dxa"/><w:end w:w="100" w:type="dxa"/>' +
          "</w:tcMar></w:tcPr>" +
          '<w:p><w:r><w:t xml:space="preserve">Cell</w:t></w:r></w:p>' +
          "</w:tc></w:tr></w:tbl>"
      )
    );
    let state = openState(opened.doc, opened.session);
    let textPos = -1;
    state.doc.descendants((node, pos) => {
      if (textPos < 0 && node.isText) textPos = pos;
      return textPos < 0;
    });
    if (textPos < 0) throw new Error("the margin document has no text");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, textPos))
    );
    state = ran(
      state,
      "pad a cell with leading and trailing margins",
      setCellPadding({ top: 6, left: 8 })
    );
    const parts = wordprocessingParts(exportDocx(state.doc, opened.session));
    expect(parts.get(opened.session.mainPartPath)).toContain(
      '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:start'
    );
    expectPartsValidate("strict cell margins", parts);
  });

  it("footnote and endnote parts validate with their main-story references", () => {
    const opened = importDocx(makeNotesDocx());
    const parts = wordprocessingParts(
      exportDocx(withEditedParagraph(opened.doc), opened.session)
    );

    expect(parts.has("word/footnotes.xml")).toBe(true);
    expect(parts.has("word/endnotes.xml")).toBe(true);
    expectPartsValidate("footnotes and endnotes", parts);
  });

  it("header and footer variants validate with their section references", () => {
    const opened = importDocx(makeHeadersFootersDocx());
    const parts = wordprocessingParts(
      exportDocx(withEditedParagraph(opened.doc), opened.session)
    );

    expect(parts.has("word/header1.xml")).toBe(true);
    expect(parts.has("word/footer1.xml")).toBe(true);
    expectPartsValidate("headers and footers", parts);
  });

  it.each(fixtureNames)(
    "%s: every WordprocessingML part validates once a paragraph was edited",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const parts = wordprocessingParts(
        exportDocx(withEditedParagraph(doc), session)
      );

      expect(parts.get(session.mainPartPath)).toContain(EDITED);
      expectPartsValidate(name, parts);
    }
  );

  /** Without this the suite would stay green on a validator that agrees to anything */
  it("turns down a body carrying an element the schema does not define", () => {
    const { doc, session } = importDocx(readFixture("kitchen-sink.docx"));
    const parts = wordprocessingParts(exportDocx(doc, session));
    const documentXml = parts.get(session.mainPartPath);
    if (documentXml === undefined) {
      throw new Error("the export wrote no body part");
    }

    const invalid = documentXml.replace("<w:body>", "<w:body><w:notInWml/>");
    expect(invalid).not.toBe(documentXml);

    const { valid, report } = validate(session.mainPartPath, invalid);
    expect(valid).toBe(false);
    expect(report).toContain(session.mainPartPath);
    expect(report).toContain("notInWml");
  });
});

describe("the exported package after an edit battery", () => {
  it.each(fixtureNames)("%s: every WordprocessingML part validates", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    expectBatteryValidates(name, doc, session);
  });

  /** The same battery over a body holding no table at all, so the one it inserts is the first */
  it.each(fixtureNames)(
    "%s: every WordprocessingML part validates with the tables dropped first",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      expectBatteryValidates(name, withoutTables(doc), session);
    }
  );
});
