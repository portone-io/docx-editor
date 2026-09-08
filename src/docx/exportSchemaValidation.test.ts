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
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { afterAll, describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  fixtureNames,
  LETTER_FIXTURE,
  LETTER_GEOMETRY,
  LETTER_SECT_PR_UNIVERSAL,
  makeDeclaredDocx,
  makeDocx,
  makeHeadersFootersDocx,
  makeNotesDocx,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
} from "../__testing__/docx";
import { posOfText } from "../__testing__/editing";
import {
  addComment,
  documentComments,
  lockSelection,
  selectionLock,
  setCommentResolved,
} from "../editor/commands";
import { toggleBulletList } from "../editor/commands/listCommands";
import { createEditorState } from "../editor/createEditor";
import { parseXml, R_NS, W_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { setCellPadding } from "../table";
import { type EditedBlock, withEditedFirst } from "./__testing__/blockEdits";
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
import { CONTENT_TYPES_PATH } from "./packageParts";
import type { SessionStore } from "./session";

const XSD_NS = "http://www.w3.org/2001/XMLSchema";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

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
  "word/commentsExtended.xml",
  "word/people.xml",
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

/**
 * A package no root of which declares more than the wordprocessing namespace: no `w14`, no `mc`,
 * no `r`. Its comments part already holds an entry, so writing a thread key into it is a splice
 * into a root that binds none of what the key needs rather than a part written from scratch.
 */
function plainlyDeclaredPackage(): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(
    makeDocx(
      '<w:p><w:commentRangeStart w:id="4"/>' +
        '<w:r><w:t xml:space="preserve">Alpha</w:t></w:r>' +
        '<w:commentRangeEnd w:id="4"/>' +
        '<w:r><w:commentReference w:id="4"/></w:r></w:p>',
      undefined,
      { prefix: "w" }
    )
  );
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId5" Target="comments.xml" ' +
      `Type="${R_NS}/comments"/></Relationships>`
  );
  parts["word/comments.xml"] = encoder.encode(
    `<w:comments xmlns:w="${W_NS}">` +
      '<w:comment w:id="4" w:author="Ada"><w:p><w:r>' +
      '<w:t xml:space="preserve">Check this</w:t></w:r></w:p></w:comment>' +
      "</w:comments>"
  );
  parts[CONTENT_TYPES_PATH] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

const EDITED = "edited before the export was validated";

/**
 * The document with the first paragraph holding text rewritten, so that the part the validator
 * reads is one the export built rather than one it handed back untouched.
 */
function withEditedParagraph(doc: PMNode): PMNode {
  return withEditedFirst(doc, "paragraph", EDITED);
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
  const snapshots = new Map<string, string>();
  const seen = new Set<string>();
  let step = 0;
  const final = afterTheBattery(
    openState(doc, session),
    (probe, before, after) => {
      const label = `${name}: ${probe.name}`;
      const previous = exportedPackage(label, before.doc, session);
      const current = exportedPackage(label, after.doc, session);
      expect(
        current.mainXml === previous.mainXml &&
          Object.keys(current.parts).every(
            (path) =>
              previous.parts[path] !== undefined &&
              bytesEqual(current.parts[path], previous.parts[path])
          ) &&
          Object.keys(current.parts).length ===
            Object.keys(previous.parts).length,
        `${label}: no exported part changed`
      ).toBe(false);
      expectEveryXmlPartParses(label, current.bytes);
      // All intermediate outputs reach xmllint. Identical parts need only one validation, and
      // batching them compiles the schema set once instead of once per command.
      for (const [path, xml] of wordprocessingParts(current.bytes)) {
        const key = `${path}\0${xml}`;
        if (seen.has(key)) continue;
        seen.add(key);
        snapshots.set(
          `${String(step).padStart(2, "0")}-${probe.name.replaceAll(" ", "_")}/${path}`,
          xml
        );
      }
      step += 1;
    }
  );
  expectPartsValidate(name, snapshots);
  expectProbesWrote(
    exportedPackage(name, final.doc, session),
    exportedPackage(name, doc, session)
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

  it.each([
    {
      name: "a prefix inherited from the grid",
      declarations: `xmlns:g="${W_NS}"`,
      revision:
        '<g:tblGridChange g:id="0"><g:tblGrid><g:gridCol g:w="900"/>' +
        "</g:tblGrid></g:tblGridChange>",
    },
    {
      name: "the default namespace inherited from the grid",
      declarations: `xmlns="${W_NS}"`,
      revision:
        '<tblGridChange w:id="0"><tblGrid><gridCol w:w="900"/>' +
        "</tblGrid></tblGridChange>",
    },
    {
      name: "a binding overridden by the revision itself",
      declarations: 'xmlns:g="urn:unused"',
      revision:
        `<g:tblGridChange xmlns:g="${W_NS}" g:id="0">` +
        '<g:tblGrid><g:gridCol g:w="900"/></g:tblGrid></g:tblGridChange>',
    },
    {
      name: "a binding overridden inside the revision",
      declarations: 'xmlns:g="urn:unused"',
      revision:
        `<w:tblGridChange w:id="0"><g:tblGrid xmlns:g="${W_NS}">` +
        '<g:gridCol g:w="900"/></g:tblGrid></w:tblGridChange>',
    },
  ])(
    "a rebuilt grid revision keeps $name",
    ({ name, declarations, revision }) => {
      const bytes = makeDocx(
        `<w:tbl><w:tblPr/><w:tblGrid ${declarations}>` +
          `<w:gridCol w:w="1000"/>${revision}</w:tblGrid>` +
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>' +
          "<w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"
      );
      expectPartsValidate(name, wordprocessingParts(bytes));
      const { doc, session } = importDocx(bytes);
      const state = createEditorState(doc);
      const edited = state.apply(
        state.tr.insertText("edited", posOfText(state.doc, "a"))
      );
      const written = exportDocx(edited.doc, session);
      const parts = wordprocessingParts(written);

      expectPartsValidate(name, parts);
      const xml = parts.get(session.mainPartPath);
      if (xml === undefined)
        throw new Error("the exported body part is missing");
      const gridChange = parseXml(xml).getElementsByTagNameNS(
        W_NS,
        "tblGridChange"
      );
      expect(gridChange).toHaveLength(1);
      const columns = gridChange[0].getElementsByTagNameNS(W_NS, "gridCol");
      expect(columns).toHaveLength(1);
      expect(columns[0].getAttributeNS(W_NS, "w")).toBe("900");
      const reopened = importDocx(written);
      expect(reopened.doc.textContent).toContain("edited");
      expect(reopened.doc.firstChild?.type.name).toBe("table");
      expect(reopened.doc.firstChild?.attrs.gridChange).toEqual(
        expect.any(String)
      );
    }
  );

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

  /**
   * A bookmark over a column stands under the row, where `CT_Row` puts `EG_RunLevelElts` between
   * its cells. It rides on the cells around it, so a rebuilt row has to put it back in a spot the
   * schema allows rather than wherever it is convenient.
   */
  it("a row-level bookmark remains valid after the row is rebuilt", () => {
    const opened = importDocx(
      makeDocx(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/>' +
          '<w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
          '<w:bookmarkStart w:id="9" w:name="Column"/>' +
          '<w:tc><w:p><w:r><w:t xml:space="preserve">First</w:t></w:r></w:p></w:tc>' +
          '<w:tc><w:p><w:r><w:t xml:space="preserve">Second</w:t></w:r></w:p></w:tc>' +
          '<w:bookmarkEnd w:id="9"/></w:tr></w:tbl>'
      )
    );
    const parts = wordprocessingParts(
      exportDocx(withEditedFirst(opened.doc, "table", EDITED), opened.session)
    );

    expect(parts.get(opened.session.mainPartPath)).toContain(
      '<w:tr><w:bookmarkStart w:id="9" w:name="Column"/><w:tc'
    );
    expectPartsValidate("row-level bookmark", parts);
  });

  it("a bookmark starting inside a cell and ending after it remains valid", () => {
    const bytes = makeDocx(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        '<w:tr><w:tc><w:p><w:bookmarkStart w:id="9" w:name="Range"/>' +
        "<w:r><w:t>First</w:t></w:r></w:p></w:tc>" +
        '<w:bookmarkEnd w:id="9"/></w:tr></w:tbl>'
    );
    expectPartsValidate(
      "original cross-cell bookmark",
      wordprocessingParts(bytes)
    );
    const opened = importDocx(bytes);
    const parts = wordprocessingParts(
      exportDocx(withEditedFirst(opened.doc, "table", EDITED), opened.session)
    );
    expectPartsValidate("rebuilt cross-cell bookmark", parts);
  });

  /**
   * The numbering part the export writes from scratch, whose root has to declare the prefix the
   * definitions inside it are written under for the schemas to read it at all.
   */
  it("a numbering part written for the first list validates", () => {
    const opened = importDocx(
      makeDeclaredDocx(
        '<w:p><w:r><w:t xml:space="preserve">First</w:t></w:r></w:p>'
      )
    );
    expect(opened.session.numberingPartPath).toBeNull();
    const state = ran(
      firstTextParagraph(openState(opened.doc, opened.session)),
      toggleBulletList
    );

    const written = exportDocx(state.doc, opened.session);
    const parts = wordprocessingParts(written);
    expect(parts.has("word/numbering.xml")).toBe(true);
    expectPartsValidate("a numbering part written from scratch", parts);
    expectEveryXmlPartParses("a numbering part written from scratch", written);
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

/** The prefixes the reports below are read in, since xmllint names a namespace by its URI */
const REPORT_PREFIXES = new Map([
  [W_NS, "w"],
  [W14_NS, "w14"],
  [MC_NS, "mc"],
]);

const SCHEMA_ERROR = "Schemas validity error : ";

/**
 * Every violation in one xmllint report, one entry per occurrence in the order the part carries
 * them, worded as xmllint worded it with the offending value left in.
 */
function violationOccurrences(report: string): readonly string[] {
  return report.split("\n").flatMap((line) => {
    const at = line.indexOf(SCHEMA_ERROR);
    if (at === -1) return [];
    return [
      line
        .slice(at + SCHEMA_ERROR.length)
        .replaceAll(
          /\{([^}]*)\}/g,
          (_whole, uri: string) => `${REPORT_PREFIXES.get(uri) ?? uri}:`
        ),
    ];
  });
}

/**
 * The same report reduced to the distinct kinds of violation in it.
 *
 * A producer writes the same invalid measurement on every cell in the document, so a pin over
 * the raw report would record how often its writer is wrong rather than what it is wrong about,
 * and one added table cell would then have to be approved as a schema change. This is the shape
 * the approved list below is written in; it says nothing about how many times, which is what
 * `violationsChanged` is for.
 */
function violationKinds(report: string): readonly string[] {
  const kinds = new Set(
    violationOccurrences(report).map((occurrence) =>
      occurrence.replace(
        /'[^']*' is not a valid value/,
        "'…' is not a valid value"
      )
    )
  );
  return Array.from(kinds).sort();
}

const strayParaId =
  "Element 'w:p', attribute 'w14:paraId': " +
  "The attribute 'w14:paraId' is not allowed.";

const notAMeasurement = (element: string, value = "…") =>
  `Element '${element}', attribute 'w:w': '${value}' is not a valid value of` +
  " the union type 'w:ST_MeasurementOrPercent'.";

/** The cell margins a producer may write as decimals, on the four sides and in a style */
const DECIMAL_CELL_MARGINS = [
  notAMeasurement("w:bottom"),
  notAMeasurement("w:left"),
  notAMeasurement("w:right"),
  notAMeasurement("w:top"),
];

/**
 * What the schemas turn down in each producer package, approved as a file a reviewer reads.
 *
 * Every one of these is markup the producer wrote and this editor hands back untouched, so the
 * pin says what a real word processor puts in a package that ECMA-376 does not describe. A diff
 * here means either that a producer file changed or that the export stopped preserving what it
 * was handed, and `__fixtures__/README.md` explains each entry under "Known gaps".
 */
const PRODUCER_VIOLATIONS: Readonly<
  Record<string, Readonly<Record<string, readonly string[]>>>
> = {
  "google-docs-export.docx": {
    "word/document.xml": [
      ...DECIMAL_CELL_MARGINS,
      strayParaId,
      "Element 'w:pgMar': The attribute 'w:gutter' is required but missing.",
      notAMeasurement("w:tblW"),
    ],
    "word/styles.xml": DECIMAL_CELL_MARGINS,
    "word/header1.xml": [strayParaId],
    "word/footer1.xml": [strayParaId],
    "word/footnotes.xml": [strayParaId],
    "word/comments.xml": [strayParaId],
  },
};

const MAIN_PART = "word/document.xml";

/**
 * What a rebuilt block stops being turned down for, per producer file and per kind of block.
 *
 * An untouched block goes out as the producer's bytes, but an edited one goes through the
 * writer, which writes the measurements it models in the schema's own form and has no model for
 * some of what a producer put in a cell. Each entry is one occurrence the untouched export
 * carries and the edited one no longer does, in the sorted order `violationsChanged` returns. A block
 * that starts dropping more of a producer's markup, or normalizing more of it, shows up here and
 * is approved rather than absorbed; `__fixtures__/README.md` explains each entry under "Known
 * gaps".
 */
const REBUILD_DROPS: Readonly<
  Record<string, Readonly<Record<EditedBlock, readonly string[]>>>
> = {
  "google-docs-export.docx": {
    paragraph: [],
    table: [
      // The paragraph of each cell continuing a vertical merge, which the writer writes empty
      `${MAIN_PART}: ${strayParaId}`,
      `${MAIN_PART}: ${strayParaId}`,
      // The width the producer wrote with a decimal point, written back as an integer
      `${MAIN_PART}: ${notAMeasurement("w:tblW", "9026.0")}`,
    ],
  },
};

function expectOnlyApprovedViolations(
  name: string,
  reports: Map<string, string>
): void {
  const found = Object.fromEntries(
    Array.from(reports, ([path, report]) => [path, violationKinds(report)])
  );
  const approved = Object.fromEntries(
    Object.entries(PRODUCER_VIOLATIONS[name] ?? {}).map(([path, kinds]) => [
      path,
      Array.from(kinds).sort(),
    ])
  );

  expect(found, `${name}: the violations the schemas found`).toEqual(approved);
}

interface ViolationsChanged {
  /** What the edited export is turned down for and the untouched one was not */
  added: string[];
  /** What the untouched export was turned down for and the edited one is not */
  dropped: string[];
}

/**
 * The violations of one export against another, occurrence by occurrence.
 *
 * The kinds of violation a producer writes are the same before and after almost any edit,
 * because the producer already wrote each kind somewhere in the file. What an edit can change is
 * how many times a part is turned down and for which values, so the two reports are compared as
 * multisets of occurrences and only the difference is returned, each entry prefixed by its part.
 */
function violationsChanged(
  untouched: Map<string, string>,
  edited: Map<string, string>
): ViolationsChanged {
  const changed: ViolationsChanged = { added: [], dropped: [] };
  const paths = new Set([...untouched.keys(), ...edited.keys()]);
  for (const path of Array.from(paths).sort()) {
    const before = violationOccurrences(untouched.get(path) ?? "");
    const after = violationOccurrences(edited.get(path) ?? "");
    const counts = new Map<string, number>();
    for (const occurrence of before) {
      counts.set(occurrence, (counts.get(occurrence) ?? 0) + 1);
    }
    for (const occurrence of after) {
      counts.set(occurrence, (counts.get(occurrence) ?? 0) - 1);
    }
    for (const [occurrence, count] of counts) {
      const into = count > 0 ? changed.dropped : changed.added;
      for (let n = 0; n < Math.abs(count); n += 1) {
        into.push(`${path}: ${occurrence}`);
      }
    }
  }
  changed.added.sort();
  changed.dropped.sort();
  return changed;
}

interface UntouchedExport {
  doc: PMNode;
  session: SessionStore;
  parts: Map<string, string>;
  reports: Map<string, string>;
}

const untouchedExports = new Map<string, UntouchedExport>();

/**
 * A producer file opened and exported with nothing edited, validated once.
 *
 * Every test below compares an edited export against this one, and compiling the schema set is
 * what a validation costs, so the untouched verdict is worked out once per file and kept.
 */
function untouchedExportOf(name: string): UntouchedExport {
  const kept = untouchedExports.get(name);
  if (kept !== undefined) return kept;
  const { doc, session } = importDocx(readProducerFixture(name));
  const parts = wordprocessingParts(exportDocx(doc, session));
  expect(parts.size).toBeGreaterThan(0);
  const opened = { doc, session, parts, reports: validateParts(parts) };
  untouchedExports.set(name, opened);
  return opened;
}

/** The edited export of a producer file, with the first block of this kind rewritten */
function editedExportOf(
  name: string,
  kind: EditedBlock
): { parts: Map<string, string>; reports: Map<string, string> } {
  const { doc, session } = untouchedExportOf(name);
  const parts = wordprocessingParts(
    exportDocx(withEditedFirst(doc, kind, EDITED), session)
  );
  expect(parts.get(session.mainPartPath)).toContain(EDITED);
  return { parts, reports: validateParts(parts) };
}

/** The document part with a page-margin element that has no gutter put inside the edited paragraph */
function withGutterlessMarginInEditedParagraph(documentXml: string): string {
  const edited = documentXml.indexOf(EDITED);
  const paragraph = documentXml.lastIndexOf("<w:p ", edited);
  const close = documentXml.indexOf("</w:pPr>", paragraph);
  if (edited === -1 || paragraph === -1 || close === -1 || close > edited) {
    throw new Error(
      "the edited paragraph has no paragraph properties to add to"
    );
  }
  return (
    documentXml.slice(0, close) +
    '<w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' +
    ' w:header="720" w:footer="720"/></w:sectPr>' +
    documentXml.slice(close)
  );
}

/** The document part with the rebuilt table's width written as a value no reader accepts */
function withUnreadableWidthOnEditedTable(documentXml: string): string {
  const edited = documentXml.indexOf(EDITED);
  const table = documentXml.lastIndexOf("<w:tbl>", edited);
  const width = documentXml.indexOf("<w:tblW ", table);
  if (edited === -1 || table === -1 || width === -1 || width > edited) {
    throw new Error("the edited table declares no width to spoil");
  }
  const end = documentXml.indexOf("/>", width);
  const spoiled = documentXml
    .slice(width, end)
    .replace(/w:w="[^"]*"/, 'w:w="NaN"');
  return documentXml.slice(0, width) + spoiled + documentXml.slice(end);
}

/**
 * The producer whose file the negative controls are run over. A control has to spoil the export
 * with a violation of a kind the producer already wrote in, so that only a comparison that counts
 * occurrences can tell the spoiled export from the approved one, and that is a property of one
 * particular file rather than of the lane.
 */
const CONTROL_FIXTURE = "google-docs-export.docx";

/**
 * The producer lane, whose files a word processor saved rather than this project.
 *
 * A hand-written fixture only ever carries markup this project chose to write, so this lane is
 * the only place the export meets what a producer actually puts in a package - and what Google
 * Docs puts in one is not valid against the transitional schemas. The export preserves those
 * bytes on every block nobody edited, so the promise the lane can hold is that the untouched
 * export is turned down for exactly what the producer wrote in, and that an edit adds no
 * occurrence of its own: not a new kind, and not one more of a kind the producer already wrote.
 */
describe("the exported producer package against the OOXML schemas", () => {
  it("there are producer fixtures, each with its violations approved", () => {
    expect(producerFixtureNames.length).toBeGreaterThan(0);
    expect(Object.keys(PRODUCER_VIOLATIONS)).toEqual(
      Array.from(producerFixtureNames)
    );
    expect(Object.keys(REBUILD_DROPS)).toEqual(
      Array.from(producerFixtureNames)
    );
    expect(producerFixtureNames).toContain(CONTROL_FIXTURE);
  });

  it.each(producerFixtureNames)(
    "%s: every WordprocessingML part validates but for the approved violations",
    (name) => {
      expectOnlyApprovedViolations(name, untouchedExportOf(name).reports);
    }
  );

  it.each(
    producerFixtureNames.flatMap((name) =>
      (["paragraph", "table"] as const).map((kind) => ({ name, kind }))
    )
  )("$name: editing a $kind adds no violation of its own", ({ name, kind }) => {
    const untouched = untouchedExportOf(name);
    const edited = editedExportOf(name, kind);

    expect(violationsChanged(untouched.reports, edited.reports)).toEqual({
      added: [],
      dropped: REBUILD_DROPS[name]?.[kind] ?? [],
    });
  });

  /**
   * The gap a comparison of kinds leaves: the producer already wrote a page margin without a
   * gutter, so a second one the edit put inside the rewritten paragraph is of an approved kind.
   */
  it("tells a second violation of an approved kind inside the edited paragraph from the first", () => {
    const untouched = untouchedExportOf(CONTROL_FIXTURE);
    const edited = editedExportOf(CONTROL_FIXTURE, "paragraph");
    const spoiled = new Map(edited.parts);
    spoiled.set(
      MAIN_PART,
      withGutterlessMarginInEditedParagraph(edited.parts.get(MAIN_PART) ?? "")
    );
    const reports = validateParts(spoiled);

    expect(violationKinds(reports.get(MAIN_PART) ?? "")).toEqual(
      violationKinds(untouched.reports.get(MAIN_PART) ?? "")
    );
    expect(violationsChanged(untouched.reports, reports).added).toEqual([
      `${MAIN_PART}: Element 'w:pgMar': The attribute 'w:gutter' is required but missing.`,
    ]);
  });

  /**
   * The same gap on the table path: a writer that spoils a width only when the original carried
   * a decimal point fires on producer input alone, and the producer already wrote a width the
   * schemas turn down, so the kind was approved before the writer went wrong.
   */
  it("tells a width the rebuilt table spoiled from the one the producer wrote", () => {
    const untouched = untouchedExportOf(CONTROL_FIXTURE);
    const edited = editedExportOf(CONTROL_FIXTURE, "table");
    const spoiled = new Map(edited.parts);
    spoiled.set(
      MAIN_PART,
      withUnreadableWidthOnEditedTable(edited.parts.get(MAIN_PART) ?? "")
    );
    const reports = validateParts(spoiled);

    expect(violationKinds(reports.get(MAIN_PART) ?? "")).toEqual(
      violationKinds(untouched.reports.get(MAIN_PART) ?? "")
    );
    expect(violationsChanged(untouched.reports, reports).added).toEqual([
      `${MAIN_PART}: ${notAMeasurement("w:tblW", "NaN")}`,
    ]);
  });
});

describe("the markup-compatibility preprocessing", () => {
  it.each([
    '<w:p mc:Ignorable="w"><w:pPr><w:jc w:val="invalid-alignment"/></w:pPr><w:r><w:t>edit</w:t></w:r></w:p>',
    '<w:p mc:Ignorable="x" mc:ProcessContent="x:wrap"><w:r><w:t>edit</w:t></w:r><x:wrap><w:r><w:notInWml/></w:r></x:wrap></w:p>',
    '<w:p><w:r><w:t>edit</w:t></w:r><mc:AlternateContent><mc:Choice Requires="w"><w:r><w:notInWml/></w:r></mc:Choice><mc:Fallback><w:r/></mc:Fallback></mc:AlternateContent></w:p>',
  ])(
    "does not erase invalid preserved WML from an edited export: %s",
    (body) => {
      const opened = importDocx(
        makeDocx(
          body.replace("<w:p", `<w:p xmlns:mc="${MC_NS}" xmlns:x="urn:unknown"`)
        )
      );
      const xml = wordprocessingParts(
        exportDocx(withEditedParagraph(opened.doc), opened.session)
      ).get(opened.session.mainPartPath);
      expect(xml).toBeDefined();
      const verdict = validate(opened.session.mainPartPath, xml ?? "");
      expect(verdict.valid, verdict.report).toBe(false);
      expect(verdict.report).toMatch(/invalid-alignment|notInWml/);
    }
  );

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
   * A package whose parts declare the wordprocessing namespace and nothing else, which is all the
   * standard asks of them: the markup a comment brings lives in namespaces of its own, and the
   * roots of the parts written have to declare each one for the schemas to read them at all.
   */
  it("accepts a comment written into a package that declares nothing else", () => {
    const opened = importDocx(plainlyDeclaredPackage());
    const state = openState(opened.doc, opened.session);
    const commented = ran(
      firstTextParagraph(state),
      addComment({
        text: "A note in a plainly declared package",
        author: "Schema test",
        initials: "ST",
        date: "2026-08-22T00:00:00Z",
      })
    );
    const added = documentComments(commented);
    const comment = added[added.length - 1];
    if (comment === undefined) throw new Error("no comment was added");
    const resolved = ran(commented, setCommentResolved(comment.id, true));

    const written = exportDocx(resolved.doc, opened.session);
    const rawComments = decode(unzipSync(written)["word/comments.xml"]);
    expect(rawComments).toContain(`<w:comments xmlns:w="${W_NS}" xmlns:w14=`);
    expect(rawComments).toContain('mc:Ignorable="w14"');

    const parts = wordprocessingParts(written);
    const commentsXml = parts.get("word/comments.xml");
    expect(commentsXml, "the export wrote no comments part").toBeDefined();
    expect(commentsXml).toContain("A note in a plainly declared package");
    expectPartsValidate("a plainly declared package", parts);
    expectEveryXmlPartParses("a plainly declared package", written);
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
  it("validates a lock added to existing content controls with properties following the lock", () => {
    const declarations = [
      '<w:date w:fullDate="2026-01-01T00:00:00Z"/>',
      '<w:comboBox><w:listItem w:value="a"/></w:comboBox>',
      "<w:text/>",
      '<w:label w:val="3"/>',
    ];
    const parts = new Map<string, string>();
    for (const [index, declaration] of declarations.entries()) {
      const { doc, session } = importDocx(
        makeDocx(
          '<w:p><w:sdt><w:sdtPr><w:id w:val="7"/>' +
            declaration +
            "</w:sdtPr><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:p>"
        )
      );
      const after = ran(
        firstTextParagraph(openState(doc, session)),
        lockSelection
      );
      expect(selectionLock(after)).toBe("locked");
      for (const [path, xml] of wordprocessingParts(
        exportDocx(after.doc, session)
      )) {
        expect(xml).toContain('w:val="sdtContentLocked"');
        parts.set(`${index}/${path}`, xml);
      }
    }
    expectPartsValidate("locked existing controls", parts);
  });

  it.each(fixtureNames)("%s: every WordprocessingML part validates", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    expectBatteryValidates(name, doc, session);
  });

  it("keeps a universal table width valid after editing a cell", () => {
    const bytes = makeDocx(
      '<w:tbl><w:tblPr><w:tblW w:w="1cm" w:type="dxa"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="1cm"/></w:tblGrid><w:tr><w:tc>' +
        '<w:tcPr><w:tcW w:w="1cm" w:type="dxa"/></w:tcPr>' +
        "<w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    );
    const { doc, session } = importDocx(bytes);
    expectPartsValidate("universal table input", wordprocessingParts(bytes));
    const edited = withEditedFirst(doc, "table", EDITED);
    const parts = wordprocessingParts(exportDocx(edited, session));
    expect(parts.get(session.mainPartPath)).toContain(EDITED);
    expectPartsValidate("edited universal table", parts);
    expect(parts.get(session.mainPartPath)).toContain('<w:gridCol w:w="567"/>');
    expect(parts.get(session.mainPartPath)).toContain(
      '<w:tblW w:w="567" w:type="dxa"/>'
    );
    expect(parts.get(session.mainPartPath)).toContain(
      '<w:tcW w:w="567" w:type="dxa"/>'
    );
  });

  /** The section remains preserved while the battery uses the paper size read from its units. */
  it("validates a document whose section is written in universal measures", () => {
    const parts = unzipSync(readFixture(LETTER_FIXTURE));
    const main = "word/document.xml";
    // The paper the fixture names, written the other way `ST_TwipsMeasure` admits it
    const rewritten = decode(parts[main]).replace(
      /<w:pgSz [^>]*\/><w:pgMar [^>]*\/>/,
      LETTER_SECT_PR_UNIVERSAL.replace(/<\/?w:sectPr>/g, "")
    );
    parts[main] = new TextEncoder().encode(rewritten);
    const { doc, session } = importDocx(zipSync(parts));

    // The same paper as `LETTER_FIXTURE`, so the battery lays out against the width it names
    expect(session.geometry).toEqual(LETTER_GEOMETRY);
    expectBatteryValidates("universal measures", doc, session);
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
