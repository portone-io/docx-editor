import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { JSDOM } from "jsdom";

const execute = promisify(execFile);
const script = fileURLToPath(
  new URL("./sanitize-fixture.mjs", import.meta.url)
);
const demo = unzipSync(
  await readFile(new URL("../__fixtures__/demo.docx", import.meta.url))
);
const instant = "2026-01-01T00:00:00Z";

async function run(t, bytes) {
  const root = await mkdtemp(join(tmpdir(), "sanitize-fixture-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const input = join(root, "input.docx");
  const output = join(root, "output.docx");
  await writeFile(input, bytes);
  return { output, result: execute(process.execPath, [script, input, output]) };
}

async function sanitized(t, xmlParts) {
  const parts = { ...demo };
  for (const [path, xml] of Object.entries(xmlParts))
    parts[path] = strToU8(xml);
  const { output, result } = await run(t, zipSync(parts));
  await result;
  return unzipSync(await readFile(output));
}

function element(xml) {
  const window = new JSDOM().window;
  try {
    const doc = new window.DOMParser().parseFromString(xml, "application/xml");
    assert.equal(doc.getElementsByTagName("parsererror").length, 0);
    return doc.documentElement;
  } finally {
    window.close();
  }
}

for (const [attribute, expected] of [
  ['w:author="Fixture O\'Example"', 'w:author="Reviewer A"'],
  ['w:initials = "FO"', 'w:initials = "RA"'],
]) {
  test(`sanitizes ${attribute} without changing the surrounding markup`, async (t) => {
    const xml = `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" ${attribute}><w:p/></w:comment></w:comments>`;
    const out = await sanitized(t, { "word/comments.xml": xml });
    assert.equal(
      strFromU8(out["word/comments.xml"]),
      xml.replace(attribute, expected)
    );
  });
}

test("re-sanitizing the committed producer fixture preserves every zip byte", async (t) => {
  const bytes = await readFile(
    new URL(
      "../__fixtures__/producers/google-docs-export.docx",
      import.meta.url
    )
  );
  const { output, result } = await run(t, bytes);
  await result;
  assert.deepEqual(await readFile(output), bytes);
});

for (const [name, bytes, error] of [
  ["non-zip input", strToU8("not a zip"), /invalid zip/i],
  ["an empty zip", zipSync({}), /no \[Content_Types\]\.xml/],
]) {
  test(`rejects ${name} with a diagnostic and no output file`, async (t) => {
    const { output, result } = await run(t, bytes);
    await assert.rejects(result, (failure) => {
      assert.equal(failure.code, 1);
      assert.match(failure.stderr, error);
      return true;
    });
    await assert.rejects(access(output), { code: "ENOENT" });
  });
}

test("anonymizes document properties and retains an empty, referenced custom-properties part", async (t) => {
  const core =
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Original Author</dc:creator><cp:lastModifiedBy>Original Editor</cp:lastModifiedBy></cp:coreProperties>';
  const app =
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Original Company</Company><Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>';
  const custom =
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Private"><vt:lpwstr>Original Value</vt:lpwstr></property></Properties>';
  const rels = strFromU8(demo["_rels/.rels"]).replace(
    "</Relationships>",
    '<Relationship Id="fixtureCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'
  );
  const types = strFromU8(demo["[Content_Types].xml"]).replace(
    "</Types>",
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'
  );
  const out = await sanitized(t, {
    "docProps/core.xml": core,
    "docProps/app.xml": app,
    "docProps/custom.xml": custom,
    "_rels/.rels": rels,
    "[Content_Types].xml": types,
  });
  const properties = element(strFromU8(out["docProps/core.xml"]));
  assert.equal(
    properties.getElementsByTagName("dc:creator")[0].textContent,
    "Fixture Author"
  );
  assert.equal(
    properties.getElementsByTagName("cp:lastModifiedBy")[0].textContent,
    "Fixture Author"
  );
  assert.equal(
    strFromU8(out["docProps/app.xml"]),
    app.replace("Original Company", "")
  );
  const emptied = element(strFromU8(out["docProps/custom.xml"]));
  assert.equal(
    emptied.namespaceURI,
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  );
  assert.equal(emptied.localName, "Properties");
  assert.equal(emptied.children.length, 0);
  assert.equal(emptied.textContent, "");
  assert.equal(strFromU8(out["_rels/.rels"]), rels);
  assert.equal(strFromU8(out["[Content_Types].xml"]), types);
});

test("pins revision and comment dates while preserving producer identifiers", async (t) => {
  const xml =
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Reviewer A" w:date="2025-12-03T12:45:00Z"><w:p w:rsidR="00112233"><w:ins w:id="8" w:author="Reviewer A" w:date=\'2025-12-04T10:00:00Z\'><w:r><w:t>text</w:t></w:r></w:ins></w:p></w:comment></w:comments>';
  const out = await sanitized(t, { "word/comments.xml": xml });
  assert.equal(
    strFromU8(out["word/comments.xml"]),
    xml.replaceAll(/2025-12-0[34]T[\d:]+Z/g, instant)
  );
});

test("removes Word account, template-path and descriptive metadata while retaining references", async (t) => {
  const core =
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Private title</dc:title><cp:keywords>Private keywords</cp:keywords><dcterms:created xsi:type="dcterms:W3CDTF">2025-09-01T00:00:00Z</dcterms:created></cp:coreProperties>';
  const app =
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Manager>Private manager</Manager><HyperlinkBase>file:///Users/private/Documents/</HyperlinkBase></Properties>';
  const people =
    '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="Private author"><w15:presenceInfo w15:providerId="AD" w15:userId="private@example.invalid"/></w15:person></w15:people>';
  const rels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="template1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///Users/private/Templates/Normal.dotm" TargetMode="External"/></Relationships>';
  const out = await sanitized(t, {
    "docProps/core.xml": core,
    "docProps/app.xml": app,
    "word/people.xml": people,
    "word/_rels/settings.xml.rels": rels,
  });
  const properties = element(strFromU8(out["docProps/core.xml"]));
  for (const name of ["dc:title", "cp:keywords"])
    assert.equal(properties.getElementsByTagName(name)[0].textContent, "");
  assert.equal(
    properties.getElementsByTagName("dcterms:created")[0].textContent,
    instant
  );
  assert.equal(
    properties
      .getElementsByTagName("dcterms:created")[0]
      .getAttribute("xsi:type"),
    "dcterms:W3CDTF"
  );
  const application = element(strFromU8(out["docProps/app.xml"]));
  for (const name of ["Manager", "HyperlinkBase"])
    assert.equal(application.getElementsByTagName(name)[0].textContent, "");
  assert.equal(
    strFromU8(out["word/people.xml"]),
    people
      .replace("Private author", "Reviewer A")
      .replace("private@example.invalid", "reviewer-a")
  );
  assert.equal(
    strFromU8(out["word/_rels/settings.xml.rels"]),
    rels.replace("file:///Users/private/Templates/", "")
  );
});
