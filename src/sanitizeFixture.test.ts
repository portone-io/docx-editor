// @vitest-environment jsdom
/**
 * The script that turns a document a word processor saved into a producer-lane fixture.
 *
 * `__fixtures__/README.md` forbids real people and authoring metadata in a fixture, and in the
 * producer lane `scripts/sanitize-fixture.mjs` is the one thing enforcing that rule, so what it
 * misses ends up committed. The script is run as a consumer runs it, over a package on disk, and
 * the packages here are shaped like what Word and Google Docs save rather than like a fixture.
 */

import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  producerFixtureNames,
  readProducerFixture,
} from "./__testing__/docx";
import { parseXml } from "./ooxml/xml";

const SCRIPT = join(
  dirname(fileURLToPath(import.meta.url)),
  "../scripts/sanitize-fixture.mjs"
);

const workDir = mkdtempSync(join(tmpdir(), "docx-editor-sanitize-"));
afterAll(() => rmSync(workDir, { recursive: true, force: true }));

interface ScriptRun {
  status: number | null;
  stderr: string;
  /** The file the script wrote, or null when it wrote none */
  output: Uint8Array | null;
}

let runs = 0;

/** The script run over these bytes on disk, the way the README says to run it */
function runScript(bytes: Uint8Array): ScriptRun {
  runs += 1;
  const input = join(workDir, `${runs}-in.docx`);
  const output = join(workDir, `${runs}-out.docx`);
  writeFileSync(input, bytes);
  const run = spawnSync(process.execPath, [SCRIPT, input, output], {
    encoding: "utf8",
  });
  if (run.error) throw run.error;
  return {
    status: run.status,
    stderr: run.stderr,
    output: existsSync(output) ? new Uint8Array(readFileSync(output)) : null,
  };
}

/** The parts of the sanitized package, decoded, after a run that has to have succeeded */
function sanitizedParts(bytes: Uint8Array): Record<string, string> {
  const run = runScript(bytes);
  expect(run.stderr).toBe("");
  expect(run.status).toBe(0);
  if (run.output === null) throw new Error("the script wrote no file");
  return Object.fromEntries(
    Object.entries(unzipSync(run.output)).map(([path, content]) => [
      path,
      decode(content),
    ])
  );
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_REL =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const CUSTOM_PROPERTIES_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const VT_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A package holding the smallest main part beside whatever parts a test adds */
function packageOf(parts: Readonly<Record<string, string>>): Uint8Array {
  const encoder = new TextEncoder();
  const all: Record<string, string> = {
    "[Content_Types].xml":
      `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NS}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>",
    "_rels/.rels":
      `${XML_DECLARATION}<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>` +
      "</Relationships>",
    "word/document.xml":
      `${XML_DECLARATION}<w:document xmlns:w="${W_NS}"><w:body>` +
      "<w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>",
    ...parts,
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(all).map(([path, xml]) => [path, encoder.encode(xml)])
    )
  );
}

/** The part a relationship target names, resolved from the part the relationships belong to */
function targetPartOf(relsPath: string, target: string): string {
  const owner = relsPath.replace(/_rels\/([^/]*)\.rels$/, "");
  const segments: string[] = [];
  for (const segment of `${owner}${target}`.split("/")) {
    if (segment === "..") segments.pop();
    else if (segment !== "" && segment !== ".") segments.push(segment);
  }
  return segments.join("/");
}

/**
 * Every internal relationship target and every content-type override that names a part the
 * package does not hold. A producer never writes one, so any found is the sanitizer's doing.
 */
function danglingReferences(parts: Readonly<Record<string, string>>): string[] {
  const dangling: string[] = [];
  for (const [path, xml] of Object.entries(parts)) {
    const document = parseXml(xml);
    if (path.endsWith(".rels")) {
      for (const relationship of Array.from(
        document.getElementsByTagNameNS(REL_NS, "Relationship")
      )) {
        if (relationship.getAttribute("TargetMode") === "External") continue;
        const target = targetPartOf(
          path,
          relationship.getAttribute("Target") ?? ""
        );
        if (!(target in parts)) dangling.push(`${path} -> ${target}`);
      }
    }
    if (path === "[Content_Types].xml") {
      for (const override of Array.from(
        document.getElementsByTagNameNS(CONTENT_TYPES_NS, "Override")
      )) {
        const part = (override.getAttribute("PartName") ?? "").slice(1);
        if (!(part in parts)) dangling.push(`${path} -> ${part}`);
      }
    }
  }
  return dangling;
}

/** The document properties Word writes, with the custom ones the author's environment added */
const WORD_PROPERTIES = {
  "[Content_Types].xml":
    `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
    "</Types>",
  "_rels/.rels":
    `${XML_DECLARATION}<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${PACKAGE_REL}/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${OFFICE_REL}/extended-properties" Target="docProps/app.xml"/>` +
    `<Relationship Id="rId4" Type="${OFFICE_REL}/custom-properties" Target="docProps/custom.xml"/>` +
    "</Relationships>",
  "docProps/core.xml":
    `${XML_DECLARATION}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:creator>Fixture O'Example</dc:creator><cp:lastModifiedBy>Fixture O'Example</cp:lastModifiedBy>" +
    "</cp:coreProperties>",
  "docProps/app.xml":
    `${XML_DECLARATION}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
    "<Application>Microsoft Office Word</Application><Company>Example Ltd</Company><AppVersion>16.0000</AppVersion>" +
    "</Properties>",
  "docProps/custom.xml":
    `${XML_DECLARATION}<Properties xmlns="${CUSTOM_PROPERTIES_NS}" xmlns:vt="${VT_NS}">` +
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ContentTypeId"><vt:lpwstr>0x0101</vt:lpwstr></property>' +
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Owner"><vt:lpwstr>Fixture O\'Example</vt:lpwstr></property>' +
    "</Properties>",
};

describe("the fixture sanitize script", () => {
  it("rewrites an author whichever quote encloses the name and however the equals sign is spaced", () => {
    const parts = sanitizedParts(
      packageOf({
        "word/comments.xml":
          `<w:comments xmlns:w="${W_NS}">` +
          `<w:comment w:author="Fixture O'Example" w:initials = "FO" w:id="0"><w:p/></w:comment>` +
          `<w:comment w:author='Says "hi"' w:initials='SH' w:id="1"><w:p/></w:comment>` +
          "</w:comments>",
        "word/people.xml":
          `<w15:people xmlns:w15="${W15_NS}">` +
          "<w15:person w15:author = 'Fixture O&apos;Example'><w15:presenceInfo w15:providerId=\"None\"/></w15:person>" +
          "</w15:people>",
      })
    );

    expect(parts["word/comments.xml"]).toContain(
      '<w:comment w:author="Reviewer A" w:initials = "RA" w:id="0">'
    );
    expect(parts["word/comments.xml"]).toContain(
      "<w:comment w:author='Reviewer A' w:initials='RA' w:id=\"1\">"
    );
    expect(parts["word/people.xml"]).toContain(
      "<w15:person w15:author = 'Reviewer A'>"
    );
    expect(JSON.stringify(parts)).not.toMatch(/Example|Says|"FO"|'SH'/);
  });

  it("keeps the custom properties part, emptied, so that nothing in the package points at a part it lacks", () => {
    const parts = sanitizedParts(packageOf(WORD_PROPERTIES));

    expect(parts["docProps/custom.xml"]).toBe(
      `${XML_DECLARATION}<Properties xmlns="${CUSTOM_PROPERTIES_NS}" xmlns:vt="${VT_NS}"/>`
    );
    expect(parts["_rels/.rels"]).toBe(WORD_PROPERTIES["_rels/.rels"]);
    expect(parts["[Content_Types].xml"]).toBe(
      WORD_PROPERTIES["[Content_Types].xml"]
    );
    expect(danglingReferences(parts)).toEqual([]);
    expect(parts["docProps/core.xml"]).toContain(
      "<dc:creator>Fixture Author</dc:creator><cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>"
    );
    expect(parts["docProps/app.xml"]).toContain(
      "<Application>Microsoft Office Word</Application><Company></Company><AppVersion>16.0000</AppVersion>"
    );
    expect(JSON.stringify(parts)).not.toMatch(/Example|Owner|0x0101/);
  });

  it("pins every date a producer stamped to the instant the zip entries carry", () => {
    const parts = sanitizedParts(
      packageOf({
        ...WORD_PROPERTIES,
        "docProps/core.xml": WORD_PROPERTIES["docProps/core.xml"].replace(
          "</cp:coreProperties>",
          '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-07T04:21:35Z</dcterms:created>' +
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-07T05:02:11Z</dcterms:modified>' +
            "</cp:coreProperties>"
        ),
        "word/document.xml":
          `${XML_DECLARATION}<w:document xmlns:w="${W_NS}"><w:body><w:p>` +
          '<w:ins w:id="1" w:author="Fixture O\'Example" w:date="2026-09-07T04:21:35Z"><w:r><w:t>added</w:t></w:r></w:ins>' +
          "<w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>",
        "word/comments.xml":
          `<w:comments xmlns:w="${W_NS}">` +
          '<w:comment w:id="0" w:author="Fixture O\'Example" w:date="2026-08-22T00:00:00Z"><w:p/></w:comment>' +
          "</w:comments>",
      })
    );

    expect(parts["word/document.xml"]).toContain(
      '<w:ins w:id="1" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z">'
    );
    expect(parts["word/comments.xml"]).toContain(
      '<w:comment w:id="0" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z">'
    );
    expect(parts["docProps/core.xml"]).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>' +
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>'
    );
    expect(JSON.stringify(parts)).not.toMatch(/2026-0[89]/);
  });

  it("refuses a zip that is not a package rather than writing an empty one", () => {
    const run = runScript(zipSync({}));

    expect(run.status).toBe(1);
    expect(run.stderr).toContain("[Content_Types].xml");
    expect(run.output).toBeNull();
  });

  it.each(producerFixtureNames)(
    "%s: is what the script gives back for itself",
    (name) => {
      const committed = readProducerFixture(name);
      const run = runScript(committed);
      expect(run.status).toBe(0);
      expect(
        run.output !== null && bytesEqual(run.output, committed),
        "the committed bytes are the sanitized bytes"
      ).toBe(true);
    }
  );
});
