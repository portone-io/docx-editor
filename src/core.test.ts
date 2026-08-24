// @vitest-environment jsdom
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { decode, readFixture } from "./__testing__/docx";
import {
  DocxImportError,
  type DocxSession,
  documentNumbering,
  documentPartPath,
  docxSchema,
  exportDocx,
  importDocx,
  type NumberingRef,
  toParagraphFormat,
} from "./core";

/** A fixture with numbered lists, so the numbering side of the entry is covered too */
const FIXTURE = "kitchen-sink.docx";

const EDITED = "edited through the core entry";

const srcDir = dirname(fileURLToPath(import.meta.url));

/**
 * Every package the entry reaches by following relative imports.
 *
 * The core entry has to stay usable without React and without a DOM-bound
 * editor view, and an accidental import is invisible until a consumer bundles
 * it. Reading the graph off disk is what makes that promise checkable.
 */
function packagesReachedBy(entry: string): string[] {
  const packages = new Set<string>();
  const visited = new Set<string>();

  const visit = (file: string): void => {
    if (visited.has(file)) return;
    visited.add(file);
    for (const [, specifier] of readFileSync(file, "utf8").matchAll(
      /from "([^"]+)"/g
    )) {
      if (!specifier.startsWith(".")) {
        packages.add(specifier);
        continue;
      }
      const path = resolve(dirname(file), specifier);
      visit(existsSync(`${path}.ts`) ? `${path}.ts` : `${path}/index.ts`);
    }
  };

  visit(entry);
  return [...packages].sort();
}

/** The body part of an exported file */
function documentXmlOf(bytes: Uint8Array, session: DocxSession): string {
  return decode(unzipSync(bytes)[documentPartPath(session)]);
}

/** Replaces the first run of text in the document, keeping the marks it carried */
function editFirstText(doc: PMNode, text: string): PMNode {
  let edited = false;
  const blocks: PMNode[] = [];
  doc.forEach((block) => {
    const inline: PMNode[] = [];
    block.forEach((child) => {
      if (!edited && child.isText) {
        inline.push(docxSchema.text(text, child.marks));
        edited = true;
      } else {
        inline.push(child);
      }
    });
    blocks.push(block.copy(Fragment.from(inline)));
  });
  if (!edited) throw new Error("the fixture has no text to edit");
  return docxSchema.nodes.doc.create(null, blocks);
}

function numberingRefsIn(doc: PMNode): NumberingRef[] {
  const refs: NumberingRef[] = [];
  doc.descendants((node) => {
    if (node.type !== docxSchema.nodes.paragraph) return true;
    const numbering = toParagraphFormat(node.attrs.format)?.numbering;
    if (numbering) refs.push(numbering);
    return false;
  });
  return refs;
}

describe("core entry", () => {
  it("reaches nothing but the zip and document-model packages", () => {
    expect(packagesReachedBy(join(srcDir, "core.ts"))).toEqual([
      "fflate",
      "prosemirror-model",
    ]);
  });

  it("carries an edit made against docxSchema back into the file", () => {
    const { doc, session } = importDocx(readFixture(FIXTURE));
    const out = exportDocx(editFirstText(doc, EDITED), session);

    expect(documentXmlOf(out, session)).toContain(EDITED);
    // the exported bytes are a docx again, and the edit survives reopening
    expect(importDocx(out).doc.textContent).toContain(EDITED);
  });

  it("resolves the list definitions its paragraphs point at", () => {
    const { doc, session } = importDocx(readFixture(FIXTURE));
    const { lists } = documentNumbering(session);

    const refs = numberingRefsIn(doc);
    expect(refs.length).toBeGreaterThan(0);
    for (const ref of refs) {
      expect(lists.get(ref.numId)?.levels.has(ref.ilvl)).toBe(true);
    }
  });

  it("refuses bytes that are not a docx", () => {
    expect(() => importDocx(new Uint8Array([1, 2, 3]))).toThrow(
      DocxImportError
    );
  });

  it("refuses a session it never handed out", () => {
    const { doc } = importDocx(readFixture(FIXTURE));
    const foreign: DocxSession = { kind: "docxSession" };

    expect(() => exportDocx(doc, foreign)).toThrow(
      "the session must be the one importDocx handed back"
    );
    expect(() => documentNumbering(foreign)).toThrow();
  });
});
