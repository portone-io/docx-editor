/**
 * Strips the authoring identity out of a DOCX a word processor saved, so the result can be
 * committed as a producer-lane fixture.
 *
 *     node scripts/sanitize-fixture.mjs in.docx out.docx
 *
 * `__fixtures__/README.md` forbids real people and authoring metadata in a fixture, and a file
 * that came out of Word, LibreOffice, or Google Docs carries the name of whoever saved it, their
 * account, the moment they saved, and paths under their profile in several places at once.
 * Everything else is left exactly as the producer wrote it: rsids and `w14:paraId` are markup the
 * producer lane exists to test, and `docProps/app.xml` keeps its `Application` and `AppVersion`
 * so the file testifies to its own origin. The README lists every rewrite, and what to do before
 * the first Word or LibreOffice file goes through.
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
  /** The account Word records for a person in `word/people.xml`, which names a user or a directory entry */
  userId: "reviewer-a",
};

/**
 * The one instant every date in a fixture is set to. A revision's `w:date` and the core
 * properties' timestamps record when somebody was editing, which is authoring metadata just as
 * the name is, so they are pinned to the same day the zip entries are stamped with.
 */
const INSTANT = "2026-01-01T00:00:00Z";

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
 * The comment, revision, and person markup that names whoever was editing and when. `w:author`
 * and `w:date` reach comments, tracked changes, and the note parts alike, so they are applied to
 * every part rather than to a list of paths that a producer might not have written.
 */
function withoutReviewerIdentity(xml) {
  const named = withAttribute(xml, "w:author", IDENTITY.reviewer);
  const initialled = withAttribute(named, "w:initials", IDENTITY.initials);
  const dated = withAttribute(initialled, "w:date", INSTANT);
  const person = withAttribute(dated, "w15:author", IDENTITY.reviewer);
  return withAttribute(person, "w15:userId", IDENTITY.userId);
}

/**
 * The template a Word document was made from, whose target is a path under the user's own
 * profile. The relationship stays, because `w:attachedTemplate` in the settings part points at
 * it, and only the file name is kept of the path.
 */
function withTemplateFileName(xml) {
  return xml.replace(
    /<Relationship\b[^>]*\/attachedTemplate"[^>]*>/g,
    (relationship) =>
      relationship.replace(
        /(\sTarget\s*=\s*)("[^"]*"|'[^']*')/,
        (_match, prefix, quoted) => {
          const quote = quoted[0];
          const fileName = quoted.slice(1, -1).split(/[\\/]/).pop();
          return `${prefix}${quote}${fileName}${quote}`;
        }
      )
  );
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
    [
      ["dc:creator", IDENTITY.author],
      ["cp:lastModifiedBy", IDENTITY.author],
      ["dcterms:created", INSTANT],
      ["dcterms:modified", INSTANT],
      // What the author typed into the document's properties, which may name anyone
      ["dc:title", ""],
      ["dc:subject", ""],
      ["dc:description", ""],
      ["cp:keywords", ""],
      ["cp:category", ""],
    ].reduce((part, [name, text]) => withElementText(part, name, text), xml),
  "docProps/app.xml": (xml) =>
    [
      ["Company", ""],
      ["Manager", ""],
      // The base every relative hyperlink resolves against, which Word fills with a local path
      ["HyperlinkBase", ""],
    ].reduce((part, [name, text]) => withElementText(part, name, text), xml),
  "docProps/custom.xml": (xml) => withoutChildren(xml, "Properties"),
};

/** Every OPC package holds this part; a zip without it is not a document to sanitize */
const CONTENT_TYPES = "[Content_Types].xml";

function sanitizePackage(bytes) {
  const parts = unzipSync(bytes);
  if (parts[CONTENT_TYPES] === undefined) {
    throw new Error(
      `the package holds no ${CONTENT_TYPES}, so it is not a DOCX`
    );
  }
  const sanitized = {};
  for (const [path, content] of Object.entries(parts)) {
    if (!path.endsWith(".xml") && !path.endsWith(".rels")) {
      sanitized[path] = content;
      continue;
    }
    const xml = decoder.decode(content);
    const properties = DOCUMENT_PROPERTIES[path];
    const rewritten = withoutReviewerIdentity(
      withTemplateFileName(properties === undefined ? xml : properties(xml))
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
  try {
    writeFileSync(output, sanitizePackage(readFileSync(input)));
  } catch (error) {
    console.error(
      `${input}: ${error instanceof Error ? error.message : error}`
    );
    process.exitCode = 1;
  }
}
