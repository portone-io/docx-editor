// @vitest-environment jsdom
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { decode, makeDocx, readFixture } from "./__testing__/docx";
import { rangeOfText } from "./__testing__/editing";
import {
  DocxImportError,
  type DocxSession,
  documentNumbering,
  documentPartPath,
  docxSchema,
  exportDocx,
  importDocx,
  type NumberingRef,
  onlyCommentsChangedBy,
  toParagraphFormat,
} from "./core";
import { addComment, updateComment } from "./editor/commands/commentCommands";
import { createEditorState } from "./editor/createEditor";

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

/**
 * The verifier a server runs over a file a browser handed back. A comment-only mode in the
 * editor is a courtesy; this is where the rule is held, so it is exercised over exported bytes
 * rather than over editor states.
 */
describe("onlyCommentsChangedBy", () => {
  const run = (text: string) =>
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const original = () => {
    const parts = unzipSync(makeDocx(`<w:p>${run("Alpha beta")}</w:p>`));
    parts["[Content_Types].xml"] = new TextEncoder().encode(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>"
    );
    return zipSync(parts);
  };

  /** The bytes after this author commented on "beta" */
  function commentedBy(authorId: string): {
    bytes: Uint8Array;
    commented: Uint8Array;
  } {
    const bytes = original();
    const { doc, session } = importDocx(bytes);
    let state = createEditorState(doc);
    const { from, to } = rangeOfText(state.doc, "beta");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to))
    );
    addComment({ text: "note", author: "Someone", authorId })(
      state,
      (tr) => (state = state.apply(tr))
    );
    return { bytes, commented: exportDocx(state.doc, session) };
  }

  it("holds for an unchanged file and for a comment the author added", () => {
    const { bytes, commented } = commentedBy("me");
    expect(onlyCommentsChangedBy(bytes, bytes, "me")).toBe(true);
    expect(onlyCommentsChangedBy(bytes, commented, "me")).toBe(true);
  });

  it("does not hold for a comment claiming another identity", () => {
    const { bytes, commented } = commentedBy("other");
    expect(onlyCommentsChangedBy(bytes, commented, "me")).toBe(false);
  });

  it("does not hold for edited text", () => {
    const bytes = original();
    const { doc, session } = importDocx(bytes);
    const edited = exportDocx(editFirstText(doc, EDITED), session);
    expect(onlyCommentsChangedBy(bytes, edited, "me")).toBe(false);
  });

  it("holds for rewriting one's own comment and not another's", () => {
    const mine = commentedBy("me").commented;
    const rewrite = (bytes: Uint8Array): Uint8Array => {
      const { doc, session } = importDocx(bytes);
      let state = createEditorState(doc, { editableComments: "all" });
      expect(
        updateComment("0", "rewritten")(
          state,
          (tr) => (state = state.apply(tr))
        )
      ).toBe(true);
      return exportDocx(state.doc, session);
    };
    const rewritten = rewrite(mine);
    expect(onlyCommentsChangedBy(mine, rewritten, "me")).toBe(true);
    expect(onlyCommentsChangedBy(mine, rewritten, "other")).toBe(false);
  });
});
