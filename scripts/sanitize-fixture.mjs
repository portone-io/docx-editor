/**
 * Strips the authoring identity out of a DOCX a word processor saved, so the result can be
 * committed as a producer-lane fixture.
 *
 *     node scripts/sanitize-fixture.mjs in.docx out.docx
 *
 * `__fixtures__/README.md` forbids real people and authoring metadata in a fixture, and a file
 * that came out of Word, LibreOffice, or Google Docs carries the name of whoever saved it in
 * several places at once. Everything else is left exactly as the producer wrote it: rsids and
 * `w14:paraId` are markup the producer lane exists to test, and `docProps/app.xml` keeps its
 * `Application` and `AppVersion` so the file testifies to its own origin.
 *
 * The repack settings match the ones `__fixtures__/README.md` prescribes for every fixture, so
 * sanitizing an already sanitized file gives back the same bytes.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { unzipSync, zipSync } from "fflate";

const IDENTITY = {
  author: "Fixture Author",
  reviewer: "Reviewer A",
  initials: "RA",
};

/** Every fixture is packed this way, so its bytes do not depend on when it was built */
const REPACK = { level: 6, mtime: new Date(2026, 0, 1) };

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

/**
 * Rewrites the value of an attribute wherever it appears, in the quotes and spacing the producer
 * used. Only the quote that opened a value closes it (XML 1.0, AttValue): a name may hold an
 * apostrophe inside double quotes, and whitespace is allowed on either side of the equals sign.
 */
function withAttribute(xml, name, value) {
  return xml.replace(
    new RegExp(`(\\s${name}\\s*=\\s*)("[^"]*"|'[^']*')`, "g"),
    (_match, prefix, quoted) => `${prefix}${quoted[0]}${value}${quoted[0]}`
  );
}

/** Rewrites an element to hold nothing, keeping its attributes, and leaves an absent one alone */
function withoutChildren(xml, name) {
  return xml.replace(
    new RegExp(`(<${name}(?:\\s[^>]*?)?)>[\\s\\S]*</${name}>`),
    "$1/>"
  );
}

/** Rewrites the text of an element that holds nothing but text, and leaves an absent one alone */
function withElementText(xml, name, text) {
  return xml.replace(
    new RegExp(
      `<${name}(\\s[^>]*)?/>|<${name}(\\s[^>]*)?>[^<]*</${name}>`,
      "g"
    ),
    (_match, selfClosing, open) =>
      `<${name}${selfClosing ?? open ?? ""}>${text}</${name}>`
  );
}

/**
 * The comment, revision, and person markup that names whoever was editing. `w:author` reaches
 * comments, tracked changes, and the note parts alike, so it is applied to every part rather
 * than to a list of paths that a producer might not have written.
 */
function withoutReviewerIdentity(xml) {
  const named = withAttribute(xml, "w:author", IDENTITY.reviewer);
  const initialled = withAttribute(named, "w:initials", IDENTITY.initials);
  return withAttribute(initialled, "w15:author", IDENTITY.reviewer);
}

/**
 * The document properties, which not every producer writes at all.
 *
 * The custom properties are kept as a part and emptied rather than deleted: the package's
 * relationships and content types name the part, and a fixture whose `_rels/.rels` points at a
 * part it does not hold would be a package no producer saved. Their names and values alike are
 * whatever the author's environment put there, so nothing of them is kept.
 */
const DOCUMENT_PROPERTIES = {
  "docProps/core.xml": (xml) =>
    withElementText(
      withElementText(xml, "dc:creator", IDENTITY.author),
      "cp:lastModifiedBy",
      IDENTITY.author
    ),
  "docProps/app.xml": (xml) => withElementText(xml, "Company", ""),
  "docProps/custom.xml": (xml) => withoutChildren(xml, "Properties"),
};

function sanitizePackage(bytes) {
  const parts = unzipSync(bytes);
  const sanitized = {};
  for (const [path, content] of Object.entries(parts)) {
    if (!path.endsWith(".xml") && !path.endsWith(".rels")) {
      sanitized[path] = content;
      continue;
    }
    const xml = decoder.decode(content);
    const properties = DOCUMENT_PROPERTIES[path];
    const rewritten = withoutReviewerIdentity(
      properties === undefined ? xml : properties(xml)
    );
    sanitized[path] = xml === rewritten ? content : encoder.encode(rewritten);
  }
  return zipSync(sanitized, REPACK);
}

const [input, output] = process.argv.slice(2);
if (input === undefined || output === undefined) {
  console.error("usage: node scripts/sanitize-fixture.mjs in.docx out.docx");
  process.exitCode = 1;
} else {
  writeFileSync(output, sanitizePackage(readFileSync(input)));
}
