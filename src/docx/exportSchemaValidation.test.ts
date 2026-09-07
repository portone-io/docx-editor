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
import { type EditorState, TextSelection } from "prosemirror-state";
import { afterAll, describe, expect, it } from "vitest";
import {
  decode,
  fixtureNames,
  makeDocx,
  makeHeadersFootersDocx,
  makeNotesDocx,
  readFixture,
} from "../__testing__/docx";
import {
  addComment,
  documentComments,
  setCommentResolved,
} from "../editor/commands";
import { parseXml, W_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { setCellPadding } from "../table";
import { withoutIgnorableMarkup } from "./__testing__/mce";
import {
  afterTheBattery,
  expectProbesWrote,
  exportedPackage,
  openState,
  ran,
} from "./__testing__/writerProbes";
import { MC_NS, W14_NS } from "./comments/constants";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import type { SessionStore } from "./session";

const XSD_NS = "http://www.w3.org/2001/XMLSchema";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const CONTENT_TYPES_PATH = "[Content_Types].xml";

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
 * The parts of an exported package that WordprocessingML describes, preprocessed as ECMA-376
 * part 3 has a consumer preprocess them and ready to validate.
 *
 * Word writes `mc:Ignorable` on the root of every part it saves, so our fixtures carry it too,
 * and the thread markup this package writes for a comment declares `w14` ignorable in the same
 * way. Reading either against the part 1 schemas without that step reports a document a
 * conforming consumer accepts.
 */
function wordprocessingParts(bytes: Uint8Array): Map<string, string> {
  const parts = new Map<string, string>();
  for (const [path, data] of Object.entries(unzipSync(bytes))) {
    if (!path.endsWith(".xml")) continue;
    const xml = decode(data);
    if (parseXml(xml).documentElement.namespaceURI === W_NS) {
      parts.set(path, withoutIgnorableMarkup(xml));
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

/**
 * Every part of a package that holds XML, whichever vocabulary it is written in.
 *
 * The schemas committed here describe the wordprocessing vocabulary alone, so `.rels`, the
 * content types, and the parts the comment writer adds beside `word/comments.xml` are validated
 * by nothing. Reading each one back is the least that can be said of them, and it is what catches
 * a writer that emitted a package no reader gets past at all.
 */
function xmlParts(bytes: Uint8Array): Map<string, string> {
  const parts = new Map<string, string>();
  for (const [path, data] of Object.entries(unzipSync(bytes))) {
    if (path.endsWith(".xml") || path.endsWith(".rels")) {
      parts.set(path, decode(data));
    }
  }
  return parts;
}

/**
 * The parts of an exported package that no committed schema describes, which the test below
 * holds the package to holding. Without the list, a package that stopped writing one of them
 * would still pass a test that only reads what it finds.
 */
const UNDESCRIBED_PARTS: readonly string[] = [
  CONTENT_TYPES_PATH,
  "_rels/.rels",
  "word/_rels/document.xml.rels",
  "word/comments.xml",
];

function expectEveryXmlPartParses(name: string, bytes: Uint8Array): void {
  const parts = xmlParts(bytes);
  expect(parts.size).toBeGreaterThan(0);
  const unreadable = Array.from(parts).flatMap(([path, xml]) => {
    try {
      parseXml(xml);
      return [];
    } catch (error) {
      return [`${name} ${path}: ${String(error)}`];
    }
  });
  expect(unreadable, `${name}: parts no reader gets past`).toEqual([]);
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

/** The same document with its tables dropped, so the one a probe inserts is the only one */
function withoutTables(doc: PMNode): PMNode {
  const blocks: PMNode[] = [];
  doc.forEach((block) => {
    if (block.type.name !== "table") blocks.push(block);
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

/** The whole of the first body paragraph that holds text, which a probe of its own works on */
function firstTextParagraph(state: EditorState): EditorState {
  let pos = 0;
  for (let index = 0; index < state.doc.childCount; index += 1) {
    const block = state.doc.child(index);
    if (block.type.name === "paragraph" && block.textContent !== "") {
      return state.apply(
        state.tr.setSelection(
          TextSelection.create(state.doc, pos + 1, pos + 1 + block.content.size)
        )
      );
    }
    pos += block.nodeSize;
  }
  throw new Error("the document has no paragraph holding text");
}

/** Runs every writer probe over one document and reads back what the export had to write for it */
function expectBatteryValidates(
  name: string,
  doc: PMNode,
  session: SessionStore
): void {
  const exported = exportedPackage(
    name,
    afterTheBattery(openState(doc, session)).doc,
    session
  );
  expectPartsValidate(name, wordprocessingParts(exported.bytes));
  expectProbesWrote(exported, exportedPackage(name, doc, session));
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
    state = ran(state, setCellPadding({ top: 6, left: 8 }));
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

describe("the markup-compatibility preprocessing", () => {
  const IGNORED_NS = W14_NS;
  const KEPT_NS = "urn:example:declared-but-not-ignorable";

  /** A body carrying one attribute from each namespace, ignorable named on the root */
  function documentWith(attributes: string): string {
    return (
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${IGNORED_NS}"` +
      ` xmlns:x="${KEPT_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14">` +
      `<w:body><w:p ${attributes}/></w:body></w:document>`
    );
  }

  /**
   * The negative control of the preprocessing: it is the declaration that decides, not the
   * namespace being foreign, so an attribute no `mc:Ignorable` covers still reaches the
   * validator and is still turned down.
   */
  it("removes an attribute an ignorable namespace holds and keeps one no declaration covers", () => {
    const both = withoutIgnorableMarkup(
      documentWith('w14:paraId="410A1204" x:kept="yes"')
    );
    expect(both).not.toContain("paraId");
    expect(both).not.toContain("Ignorable");
    expect(both).toContain('x:kept="yes"');

    const rejected = validate("word/document.xml", both);
    expect(rejected.valid, rejected.report).toBe(false);
    expect(rejected.report).toContain("kept");

    const ignorableOnly = withoutIgnorableMarkup(
      documentWith('w14:paraId="410A1204"')
    );
    const accepted = validate("word/document.xml", ignorableOnly);
    expect(accepted.valid, accepted.report).toBe(true);
  });

  /**
   * Part 3 has the consumer that understands none of the alternatives read the fallback, so the
   * markup validated is the markup such a consumer sees. Word writes `mc:AlternateContent` around
   * a shape it draws two ways, and leaving it standing would turn down every document that
   * carries one rather than reporting anything about this package's own writing.
   */
  it("reads an alternate-content block as the fallback a part 1 consumer takes", () => {
    const processed = withoutIgnorableMarkup(
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${IGNORED_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14">` +
        "<w:body><w:p><w:r><mc:AlternateContent>" +
        '<mc:Choice Requires="w14"><w:t xml:space="preserve">the choice</w:t></mc:Choice>' +
        '<mc:Fallback><w:t xml:space="preserve">the fallback</w:t></mc:Fallback>' +
        "</mc:AlternateContent></w:r></w:p></w:body></w:document>"
    );

    expect(processed).toContain("the fallback");
    expect(processed).not.toContain("the choice");
    expect(processed).not.toContain("AlternateContent");

    const { valid, report } = validate("word/document.xml", processed);
    expect(valid, report).toBe(true);
  });

  /**
   * The comment writer declares `w14` ignorable and hangs the thread key off `w14:paraId`
   * (`docx/comments/writing`), which is markup this suite read against the part 1 schemas until
   * the preprocessing above arrived. Every command that resolves, replies to or edits a thread
   * writes it, so this stands for all of them.
   */
  it("accepts a resolved thread the editor wrote", () => {
    const { doc, session } = importDocx(readFixture("kitchen-sink.docx"));
    const state = openState(doc, session);
    const commented = ran(
      firstTextParagraph(state),
      addComment({
        text: "The comment whose thread is resolved",
        author: "Schema test",
        initials: "ST",
        date: "2026-08-22T00:00:00Z",
      })
    );
    const added = documentComments(commented);
    const comment = added[added.length - 1];
    if (comment === undefined) throw new Error("no comment was added");
    const resolved = ran(commented, setCommentResolved(comment.id, true));

    const parts = wordprocessingParts(exportDocx(resolved.doc, session));
    const commentsXml = parts.get("word/comments.xml");
    expect(commentsXml, "the export wrote no comments part").toBeDefined();
    expect(commentsXml).toContain("The comment whose thread is resolved");
    expectPartsValidate("a resolved thread", parts);
  });
});

describe("the exported package after an edit battery", () => {
  it.each(fixtureNames)("%s: every WordprocessingML part validates", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    expectBatteryValidates(name, doc, session);
  });

  /**
   * The parts no committed schema describes. The battery is what puts most of them in the
   * package: a comment brings `word/comments.xml` and the thread part beside it, an image brings
   * a media relationship, and every one of them is named in the content types.
   */
  it.each(fixtureNames)(
    "%s: every XML part of the exported package parses",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const bytes = exportDocx(
        afterTheBattery(openState(doc, session)).doc,
        session
      );

      const parts = xmlParts(bytes);
      for (const path of UNDESCRIBED_PARTS) {
        expect(parts.has(path), `${name} wrote no ${path}`).toBe(true);
      }
      expectEveryXmlPartParses(name, bytes);
    }
  );

  /** The same battery over a body holding no table at all, so the one it inserts is the first */
  it.each(fixtureNames)(
    "%s: every WordprocessingML part validates with the tables dropped first",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      expectBatteryValidates(name, withoutTables(doc), session);
    }
  );
});
