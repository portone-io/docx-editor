// @vitest-environment jsdom
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  decode,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
} from "./__testing__/docx";
import { rangeOfText } from "./__testing__/editing";
import {
  type CommentOnlyVerdict,
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
import { isCommentNode } from "./schema/protection";

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
 * rather than over editor states, and over the whole package rather than the story alone.
 */
describe("onlyCommentsChangedBy", () => {
  const encoder = new TextEncoder();
  const run = (text: string) =>
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const STYLES_PART = "word/styles.xml";
  const DOCUMENT_PART = "word/document.xml";
  const DOCUMENT_RELS_PART = "word/_rels/document.xml.rels";

  const original = () => {
    const parts = unzipSync(
      makeDocx(
        `<w:p>${run("Alpha beta")}</w:p><w:p>${run("Gamma")}</w:p>` +
          LETTER_SECT_PR,
        '<w:sz w:val="20"/>'
      )
    );
    parts["[Content_Types].xml"] = encoder.encode(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>"
    );
    return zipSync(parts);
  };

  /** The package with one part written over, which is every shape a submission can be tampered in */
  function repacked(
    bytes: Uint8Array,
    changes: Readonly<Record<string, string>>
  ): Uint8Array {
    const parts = unzipSync(bytes);
    for (const [path, text] of Object.entries(changes)) {
      parts[path] = encoder.encode(text);
    }
    return zipSync(parts);
  }

  function partText(bytes: Uint8Array, path: string): string {
    return decode(unzipSync(bytes)[path]);
  }

  /** The bytes before and after this author commented on "beta" */
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

  /** Another author's comment, with the people part rewritten to claim it for "me" */
  function identityStolen(): { theirs: Uint8Array; stolen: Uint8Array } {
    const theirs = commentedBy("other").commented;
    const peoplePart = Object.keys(unzipSync(theirs)).find((path) =>
      path.endsWith("people.xml")
    );
    if (peoplePart === undefined) throw new Error("no people part");
    const stolen = repacked(theirs, {
      [peoplePart]: partText(theirs, peoplePart).replace(
        'userId="other"',
        'userId="me"'
      ),
    });
    return { theirs, stolen };
  }

  const allowed: CommentOnlyVerdict = { ok: true };
  const refusedFor = (
    reason: "body-changed" | "comment-not-owned" | "comment-author-forged"
  ): CommentOnlyVerdict => ({ ok: false, reason });
  const partRefused = (part: string): CommentOnlyVerdict => ({
    ok: false,
    reason: "part-changed",
    part,
  });

  describe("over the document story", () => {
    it("holds for an unchanged file and for a comment the author added", () => {
      const { bytes, commented } = commentedBy("me");
      expect(onlyCommentsChangedBy(bytes, bytes, "me")).toEqual(allowed);
      expect(onlyCommentsChangedBy(bytes, commented, "me")).toEqual(allowed);
    });

    it("does not hold for a comment claiming another identity", () => {
      const { bytes, commented } = commentedBy("other");
      expect(onlyCommentsChangedBy(bytes, commented, "me")).toEqual(
        refusedFor("comment-author-forged")
      );
    });

    it("does not hold for edited text", () => {
      const bytes = original();
      const { doc, session } = importDocx(bytes);
      const edited = exportDocx(editFirstText(doc, EDITED), session);
      expect(onlyCommentsChangedBy(bytes, edited, "me")).toEqual(
        refusedFor("body-changed")
      );
    });

    it("holds for rewriting one's own comment and not another's", () => {
      const mine = commentedBy("me").commented;
      const { doc, session } = importDocx(mine);
      let state = createEditorState(doc, { editableComments: "all" });
      expect(
        updateComment("0", "rewritten")(
          state,
          (tr) => (state = state.apply(tr))
        )
      ).toBe(true);
      const rewritten = exportDocx(state.doc, session);
      expect(onlyCommentsChangedBy(mine, rewritten, "me")).toEqual(allowed);
      expect(onlyCommentsChangedBy(mine, rewritten, "other")).toEqual(
        refusedFor("comment-not-owned")
      );
    });

    it("does not hold for a comment carried onto other text by anyone else", () => {
      const mine = commentedBy("me").commented;
      const { doc, session } = importDocx(mine);
      const state = createEditorState(doc, { editableComments: "all" });
      const markers: { pos: number; node: PMNode }[] = [];
      state.doc.descendants((node, pos) => {
        if (isCommentNode(node)) markers.push({ pos, node });
        return true;
      });
      const tr = state.tr;
      for (const { pos, node } of [...markers].reverse()) {
        tr.delete(pos, pos + node.nodeSize);
      }
      const target = rangeOfText(tr.doc, "Gamma");
      const [start, end, reference] = markers.map((marker) => marker.node);
      tr.insert(target.to, [end, reference]);
      tr.insert(target.from, start);
      const moved = exportDocx(tr.doc, session);
      expect(onlyCommentsChangedBy(mine, moved, "me")).toEqual(allowed);
      expect(onlyCommentsChangedBy(mine, moved, "other")).toEqual(
        refusedFor("comment-not-owned")
      );
    });

    it("does not hold for a people part rewritten to claim another's comment", () => {
      const { theirs, stolen } = identityStolen();
      expect(onlyCommentsChangedBy(theirs, stolen, "me")).toEqual(
        refusedFor("comment-author-forged")
      );
    });
  });

  describe("for a moderator's file", () => {
    it("holds for another author's comment rewritten under `all`", () => {
      const mine = commentedBy("me").commented;
      const { doc, session } = importDocx(mine);
      let state = createEditorState(doc, { editableComments: "all" });
      updateComment("0", "rewritten")(state, (tr) => (state = state.apply(tr)));
      const rewritten = exportDocx(state.doc, session);
      expect(onlyCommentsChangedBy(mine, rewritten, "other")).toEqual(
        refusedFor("comment-not-owned")
      );
      expect(
        onlyCommentsChangedBy(mine, rewritten, "other", {
          editableComments: "all",
        })
      ).toEqual(allowed);
    });

    it("does not hold for a rewritten identity under `all` either", () => {
      const { theirs, stolen } = identityStolen();
      expect(
        onlyCommentsChangedBy(theirs, stolen, "me", { editableComments: "all" })
      ).toEqual(refusedFor("comment-author-forged"));
    });
  });

  describe("over the rest of the package", () => {
    it("does not hold for a part the submission added", () => {
      const bytes = original();
      const withHeader = repacked(bytes, {
        "word/header1.xml": "<w:hdr/>",
      });
      expect(onlyCommentsChangedBy(bytes, withHeader, "me")).toEqual(
        partRefused("word/header1.xml")
      );
    });

    it("does not hold for a part the submission took away", () => {
      const bytes = original();
      const parts = unzipSync(bytes);
      delete parts[STYLES_PART];
      expect(onlyCommentsChangedBy(bytes, zipSync(parts), "me")).toEqual(
        partRefused(STYLES_PART)
      );
    });

    it("does not hold for a rewritten styles part", () => {
      const bytes = original();
      const restyled = repacked(bytes, {
        [STYLES_PART]: partText(bytes, STYLES_PART).replace(
          'w:val="20"',
          'w:val="48"'
        ),
      });
      expect(onlyCommentsChangedBy(bytes, restyled, "me")).toEqual(
        partRefused(STYLES_PART)
      );
    });

    it("does not hold for section properties the story never carried", () => {
      const bytes = original();
      const remargined = repacked(bytes, {
        [DOCUMENT_PART]: partText(bytes, DOCUMENT_PART).replace(
          'w:left="1440"',
          'w:left="720"'
        ),
      });
      expect(onlyCommentsChangedBy(bytes, remargined, "me")).toEqual(
        partRefused(DOCUMENT_PART)
      );
    });

    it("does not hold for a relationship the submission added", () => {
      const bytes = original();
      const related = repacked(bytes, {
        [DOCUMENT_RELS_PART]: partText(bytes, DOCUMENT_RELS_PART).replace(
          "</Relationships>",
          '<Relationship Id="rId9" Target="https://example.com"' +
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"' +
            ' TargetMode="External"/></Relationships>'
        ),
      });
      expect(onlyCommentsChangedBy(bytes, related, "me")).toEqual({
        ok: false,
        reason: "relationship-changed",
        part: DOCUMENT_RELS_PART,
      });
    });

    it("holds for the parts a comment of one's own is written across", () => {
      const { bytes, commented } = commentedBy("me");
      const added = Object.keys(unzipSync(commented)).filter(
        (path) => !(path in unzipSync(bytes))
      );
      expect(added).toContain("word/comments.xml");
      expect(onlyCommentsChangedBy(bytes, commented, "me")).toEqual(allowed);
    });

    it("turns bytes that are not a docx down the way opening one does", () => {
      expect(() =>
        onlyCommentsChangedBy(original(), new Uint8Array([1, 2, 3]), "me")
      ).toThrow(DocxImportError);
    });
  });
});
