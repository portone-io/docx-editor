// @vitest-environment jsdom
/**
 * The producer lane: what happens when the importer meets a document a word processor saved.
 *
 * `__fixtures__/` is otherwise built from controlled XML, which means the corpus only ever
 * exercises markup this project chose to write. A real producer writes rsids on every run,
 * `w14:paraId` on every paragraph, revision markup, machine-generated bookmark names, and
 * measurements in shapes the schema authors did not expect. These three tests are the promises
 * that matter most on such a file: it opens, an untouched round trip gives the package back
 * unchanged, and one edited paragraph leaves every other byte alone.
 *
 * `__fixtures__/README.md` records where each file came from and what it is known not to keep.
 */

import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  producerFixtureNames,
  readProducerFixture,
  surroundings,
} from "../__testing__/docx";
import { createEditorState } from "../editor/createEditor";
import { parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

const EDIT = "edited by the producer lane";

/** The first body paragraph holding text, which is the one an editing session reaches first */
function firstEditableParagraph(doc: PMNode): number {
  let index = -1;
  doc.forEach((block, _offset, at) => {
    if (index !== -1) return;
    if (block.type.name === "paragraph" && block.textContent !== "") index = at;
  });
  if (index === -1) throw new Error("the fixture has no paragraph to edit");
  return index;
}

function withEditedText(paragraph: PMNode): PMNode {
  const inline: PMNode[] = [];
  let edited = false;
  paragraph.forEach((child) => {
    if (!edited && child.isText) {
      inline.push(docxSchema.text(EDIT, child.marks));
      edited = true;
    } else {
      inline.push(child);
    }
  });
  return paragraph.copy(Fragment.from(inline));
}

function withBlockReplaced(doc: PMNode, index: number): PMNode {
  const blocks: PMNode[] = [];
  doc.forEach((block, _offset, at) => {
    blocks.push(at === index ? withEditedText(block) : block);
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

describe("a document a word processor saved", () => {
  it("there are producer fixtures", () => {
    expect(producerFixtureNames.length).toBeGreaterThan(0);
  });

  it.each(producerFixtureNames)("%s: opens without a refusal", (name) => {
    const { doc, session } = importDocx(readProducerFixture(name));

    expect(doc.childCount).toBeGreaterThan(0);
    expect(session.blocks.length).toBe(doc.childCount);
  });

  it.each(producerFixtureNames)(
    "%s: exports every part byte for byte when nothing was edited",
    (name) => {
      const bytes = readProducerFixture(name);
      const { doc, session } = importDocx(bytes);

      // The path a real session takes: the display attrs are worked out again over the import
      const out = exportDocx(createEditorState(doc).doc, session);

      const original = unzipSync(bytes);
      const exported = unzipSync(out);
      expect(Object.keys(exported)).toEqual(Object.keys(original));
      for (const path of Object.keys(original)) {
        if (path === session.mainPartPath) {
          expect(decode(exported[path])).toBe(decode(original[path]));
        }
        expect(bytesEqual(exported[path], original[path]), path).toBe(true);
      }
    }
  );

  it.each(producerFixtureNames)(
    "%s: editing the first editable paragraph changes only that block's bytes",
    (name) => {
      const bytes = readProducerFixture(name);
      const { doc, session } = importDocx(bytes);
      const index = firstEditableParagraph(doc);

      const out = exportDocx(withBlockReplaced(doc, index), session);
      const exported = unzipSync(out);
      const documentXml = decode(exported[session.mainPartPath]);

      const { head, tail } = surroundings(session, index);
      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);

      const rebuilt = documentXml.slice(
        head.length,
        documentXml.length - tail.length
      );
      expect(rebuilt).toContain(EDIT);
      // A rebuilt block declares no prefix the producer had not declared on the block itself
      expect(namespaceDecls(rebuilt)).toEqual(
        namespaceDecls(session.blocks[index].xml)
      );
      expect(() => parseXml(wrapped(rebuilt))).not.toThrow();

      const original = unzipSync(bytes);
      for (const path of Object.keys(original)) {
        if (path === session.mainPartPath) continue;
        expect(bytesEqual(exported[path], original[path]), path).toBe(true);
      }

      expect(importDocx(out).doc.textContent).toContain(EDIT);
    }
  );
});

function namespaceDecls(xml: string): string[] {
  return xml.match(/xmlns(:[\w-]+)?="[^"]*"/g) ?? [];
}

/** A block fragment carries no declarations of its own, so parsing it needs a root that does */
function wrapped(fragment: string): string {
  return (
    '<w:wrap xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    fragment +
    "</w:wrap>"
  );
}
