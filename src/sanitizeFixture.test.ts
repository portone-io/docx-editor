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
