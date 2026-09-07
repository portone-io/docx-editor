// @vitest-environment jsdom
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  decode,
  fixtureNames,
  importErrorCode,
  LETTER_SECT_PR,
  makeDocx,
  ONE_LIST_NUMBERING,
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
  parseNumbering,
  toParagraphFormat,
  type XmlParser,
} from "./core";
import { COMMENTS_REL_TYPE, PEOPLE_REL_TYPE } from "./docx/comments/constants";
import {
  addComment,
  canAddComment,
  setCommentResolved,
  updateComment,
} from "./editor/commands/commentCommands";
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
  it("constructs one global parser for all the XML an import reads", () => {
    let constructed = 0;
    class CountingParser extends DOMParser {
      constructor() {
        super();
        constructed += 1;
      }
    }
    vi.stubGlobal("DOMParser", CountingParser);

    try {
      const { doc } = importDocx(readFixture("demo.docx"));

      expect(doc.textContent).not.toBe("");
      expect(constructed).toBe(1);
    } finally {
      vi.unstubAllGlobals();
    }
  });

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

  /**
   * A comment is legitimate wherever a paragraph is, so the verdict has to hold over every
   * paragraph of the corpus rather than over the two a hand-written case reaches. A paragraph
   * inside a table is the one a comparison over the model turned down, since a commented table is
   * rebuilt and comes back worded the way this editor words it.
   */
  describe("over every paragraph of every fixture", () => {
    interface Spot {
      from: number;
      to: number;
      text: string;
    }

    /** The first character of every paragraph carrying text, as a range a comment can be put over */
    function everyParagraph(doc: PMNode): Spot[] {
      const spots: Spot[] = [];
      doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.textContent.length > 0) {
          spots.push({
            from: pos + 1,
            to: pos + 2,
            text: node.textContent.slice(0, 40),
          });
        }
        return node.type.name !== "paragraph";
      });
      return spots;
    }

    /**
     * The paragraphs that take no comment, which are the ones a locked content control holds.
     *
     * Named rather than counted, so a change that starts refusing comments elsewhere fails here
     * instead of quietly shrinking what the sweep below covers.
     */
    const LOCKED: Readonly<Record<string, readonly string[]>> = {
      "kitchen-sink.docx": ["Settled"],
    };

    const SWEEP_TIMEOUT_MS = 120_000;

    it.each(fixtureNames)(
      "holds for a comment of one's own in every paragraph of %s, table cells included",
      (name) => {
        const bytes = readFixture(name);
        const { doc, session } = importDocx(bytes);
        const spots = everyParagraph(doc);
        expect(spots.length).toBeGreaterThan(0);

        const refused: string[] = [];
        for (const { from, to, text } of spots) {
          // A state of its own for each, so every submission differs in that one comment alone
          let state = createEditorState(doc);
          state = state.apply(
            state.tr.setSelection(TextSelection.create(state.doc, from, to))
          );
          if (!canAddComment(state)) {
            refused.push(text);
            continue;
          }
          const before = state.doc;
          expect(
            addComment({ text: "note", author: "Someone", authorId: "me" })(
              state,
              (tr) => (state = state.apply(tr))
            )
          ).toBe(true);
          expect(state.doc.eq(before)).toBe(false);
          expect(
            onlyCommentsChangedBy(bytes, exportDocx(state.doc, session), "me")
          ).toEqual(allowed);
        }
        expect(refused).toEqual(LOCKED[name] ?? []);
      },
      // A comment on every paragraph of a real document, each one exported and read back twice by
      // the verifier. Minutes of work on a slow runner, and the corpus is the point of the case
      SWEEP_TIMEOUT_MS
    );

    it("still does not hold for a cell whose text was typed into", () => {
      const bytes = readFixture(FIXTURE);
      const { doc, session } = importDocx(bytes);
      let state = createEditorState(doc);
      const inACell = everyParagraph(state.doc).find(
        (spot) =>
          state.doc.resolve(spot.from).node(-1).type.name === "tableCell"
      );
      if (inACell === undefined)
        throw new Error("the fixture has no table text");
      state = state.apply(
        state.tr.insertText(EDITED, inACell.from, inACell.to)
      );

      expect(
        onlyCommentsChangedBy(bytes, exportDocx(state.doc, session), "me")
      ).toEqual(refusedFor("body-changed"));
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

    /**
     * The excuse the three comment parts get is an excuse for those parts, not for whichever part
     * a submission decides to relate under a comment type.
     */
    describe("for a part related as a comment part", () => {
      const COMMENTS_REL = COMMENTS_REL_TYPE;
      const PEOPLE_REL = PEOPLE_REL_TYPE;
      const relationshipsRefused: CommentOnlyVerdict = {
        ok: false,
        reason: "relationship-changed",
        part: DOCUMENT_RELS_PART,
      };

      /** The submission's relationships with one more pointing where it says */
      function alsoRelated(
        bytes: Uint8Array,
        type: string,
        target: string
      ): Uint8Array {
        return repacked(bytes, {
          [DOCUMENT_RELS_PART]: partText(bytes, DOCUMENT_RELS_PART).replace(
            "</Relationships>",
            `<Relationship Id="rId77" Type="${type}" Target="${target}"/>` +
              "</Relationships>"
          ),
        });
      }

      it("does not hold for a styles part the submission related as a second comments part", () => {
        const { bytes, commented } = commentedBy("me");
        const related = alsoRelated(commented, COMMENTS_REL, "styles.xml");
        const restyled = repacked(related, {
          [STYLES_PART]: partText(related, STYLES_PART).replace(
            'w:val="20"',
            'w:val="48"'
          ),
        });
        expect(onlyCommentsChangedBy(bytes, restyled, "me")).toEqual(
          partRefused(STYLES_PART)
        );
      });

      it("does not hold for a second relationship of a comment type", () => {
        const { bytes, commented } = commentedBy("me");
        const twice = alsoRelated(commented, PEOPLE_REL, "styles.xml");
        expect(onlyCommentsChangedBy(bytes, twice, "me")).toEqual(
          relationshipsRefused
        );
      });

      it("does not hold for a comment type the file already related", () => {
        const { commented } = commentedBy("me");
        const twice = alsoRelated(commented, COMMENTS_REL, "comments.xml");
        expect(onlyCommentsChangedBy(commented, twice, "me")).toEqual(
          relationshipsRefused
        );
      });

      it("compares a part a second comment relationship names, rather than excusing it", () => {
        const { commented } = commentedBy("me");
        const twice = alsoRelated(commented, COMMENTS_REL, "styles.xml");
        const restyled = repacked(twice, {
          [STYLES_PART]: partText(twice, STYLES_PART).replace(
            'w:val="20"',
            'w:val="48"'
          ),
        });
        expect(onlyCommentsChangedBy(commented, restyled, "me")).toEqual(
          partRefused(STYLES_PART)
        );
      });

      /**
       * The one rule the two above do not reach: a comment type the file has no relationship for
       * is a type a submission may relate, and the part it relates has to be one it brought.
       */
      it("does not hold for a part the file already had, related as its first people part", () => {
        const bytes = original();
        const related = alsoRelated(bytes, PEOPLE_REL, "styles.xml");
        const restyled = repacked(related, {
          [STYLES_PART]: partText(related, STYLES_PART).replace(
            'w:val="20"',
            'w:val="48"'
          ),
        });
        expect(onlyCommentsChangedBy(bytes, restyled, "me")).toEqual(
          relationshipsRefused
        );
      });

      it("holds for a first comment in a file relating a comment type outside the package", () => {
        const bytes = repacked(original(), {
          [DOCUMENT_RELS_PART]: partText(
            original(),
            DOCUMENT_RELS_PART
          ).replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${COMMENTS_REL}"` +
              ' Target="http://example.com/c.xml" TargetMode="External"/>' +
              "</Relationships>"
          ),
        });
        const { doc, session } = importDocx(bytes);
        let state = createEditorState(doc);
        const { from, to } = rangeOfText(state.doc, "beta");
        state = state.apply(
          state.tr.setSelection(TextSelection.create(state.doc, from, to))
        );
        addComment({ text: "note", author: "Someone", authorId: "me" })(
          state,
          (tr) => (state = state.apply(tr))
        );

        expect(
          onlyCommentsChangedBy(bytes, exportDocx(state.doc, session), "me")
        ).toEqual(allowed);
      });

      /** The extended part a file relates is the one a first comment is written into */
      it("holds for a first comment in a file that already related an extended comments part", () => {
        const EXTENDED_REL =
          "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
        const plain = original();
        const related = repacked(plain, {
          [DOCUMENT_RELS_PART]: partText(plain, DOCUMENT_RELS_PART).replace(
            "</Relationships>",
            `<Relationship Id="rId8" Type="${EXTENDED_REL}"` +
              ' Target="commentsExtended.xml"/></Relationships>'
          ),
          "word/commentsExtended.xml":
            '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
        });
        const { doc, session } = importDocx(related);
        let state = createEditorState(doc);
        const { from, to } = rangeOfText(state.doc, "beta");
        state = state.apply(
          state.tr.setSelection(TextSelection.create(state.doc, from, to))
        );
        addComment({ text: "note", author: "Someone", authorId: "me" })(
          state,
          (tr) => (state = state.apply(tr))
        );

        expect(
          onlyCommentsChangedBy(related, exportDocx(state.doc, session), "me")
        ).toEqual(allowed);

        // Settling the thread in the same session writes that part again rather than a second one
        const settled = state;
        let id = "";
        settled.doc.descendants((node) => {
          if (id === "" && node.type.name === "commentReference") {
            id = String(node.attrs.id);
          }
          return true;
        });
        let after = settled;
        expect(
          setCommentResolved(id, true)(settled, (tr) => {
            after = after.apply(tr);
          })
        ).toBe(true);

        expect(
          onlyCommentsChangedBy(related, exportDocx(after.doc, session), "me")
        ).toEqual(allowed);
      });

      it("does not hold for a comment relationship the submission points outside the package", () => {
        const { bytes, commented } = commentedBy("me");
        const outward = repacked(commented, {
          [DOCUMENT_RELS_PART]: partText(commented, DOCUMENT_RELS_PART).replace(
            "</Relationships>",
            `<Relationship Id="rId88" Type="${PEOPLE_REL}"` +
              ' Target="http://example.com/p.xml" TargetMode="External"/>' +
              "</Relationships>"
          ),
        });
        expect(onlyCommentsChangedBy(bytes, outward, "me")).toEqual(
          relationshipsRefused
        );
      });

      /**
       * Two relationships under one id are read differently depending on which of the two a
       * reader keeps, and the decoy is neither gained nor a change to what was there.
       */
      it("does not hold for a relationship part naming one id twice", () => {
        const { commented } = commentedBy("me");
        const rels = partText(commented, DOCUMENT_RELS_PART);
        const reused = /Id="([^"]+)"/.exec(rels)?.[1];
        const decoyed = repacked(commented, {
          [DOCUMENT_RELS_PART]: rels.replace(
            "<Relationship Id=",
            `<Relationship Id="${reused}" Type="${COMMENTS_REL}" Target="decoy.xml"/>` +
              "<Relationship Id="
          ),
        });
        expect(onlyCommentsChangedBy(commented, decoyed, "me")).toEqual(
          relationshipsRefused
        );
      });

      it("does not hold for a forged relationship written ahead of the real one", () => {
        const { bytes, commented } = commentedBy("me");
        const ahead = repacked(commented, {
          [DOCUMENT_RELS_PART]: partText(commented, DOCUMENT_RELS_PART).replace(
            "<Relationship Id=",
            `<Relationship Id="rId77" Type="${COMMENTS_REL}" Target="styles.xml"/>` +
              "<Relationship Id="
          ),
        });
        // The reader opens the forged part as the comments part, which leaves the real one
        // outside the excused set rather than putting the forged one inside it
        expect(onlyCommentsChangedBy(bytes, ahead, "me")).toEqual(
          partRefused("word/comments.xml")
        );
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

  /**
   * The three parts a comment is written across are the ones a comment edit rewrites, so the
   * package comparison passes over their bytes. What a submission put in them is read entry by
   * entry instead (`docx/comments/verifying`), or a file could carry anything at all inside a
   * comment and be answered as one where only comments changed.
   */
  describe("over the comment parts", () => {
    const COMMENTS_PART = "word/comments.xml";

    /** The same file with its comments part rewritten */
    function tampered(
      commented: Uint8Array,
      rewrite: (xml: string) => string
    ): Uint8Array {
      return repacked(commented, {
        [COMMENTS_PART]: rewrite(partText(commented, COMMENTS_PART)),
      });
    }

    it("does not hold for a field injected into a comment body", () => {
      const { bytes, commented } = commentedBy("me");
      const forged = tampered(commented, (xml) =>
        xml.replace(
          "</w:p></w:comment>",
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
            '<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "http://evil.example/x.png" </w:instrText></w:r>' +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:comment>'
        )
      );

      expect(onlyCommentsChangedBy(bytes, forged, "me")).toEqual(
        partRefused(COMMENTS_PART)
      );
    });

    it("does not hold for an orphan comment the submission added", () => {
      const { bytes, commented } = commentedBy("me");
      const orphaned = tampered(commented, (xml) =>
        xml.replace(
          "</w:comments>",
          '<w:comment w:id="999" w:author="Someone Else" w:date="2020-01-01T00:00:00Z">' +
            `<w:p>${run("ghost")}</w:p></w:comment></w:comments>`
        )
      );

      expect(onlyCommentsChangedBy(bytes, orphaned, "me")).toEqual(
        partRefused(COMMENTS_PART)
      );
    });

    it("does not hold for mc:AlternateContent inside one's own comment", () => {
      const { bytes, commented } = commentedBy("me");
      const wrapped = tampered(commented, (xml) =>
        xml.replace(
          "</w:p></w:comment>",
          '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
            `<mc:Fallback>${run("x")}</mc:Fallback></mc:AlternateContent></w:p></w:comment>`
        )
      );

      expect(onlyCommentsChangedBy(bytes, wrapped, "me")).toEqual(
        partRefused(COMMENTS_PART)
      );
    });
  });
});

/**
 * The entry points read where a server reads them, holding the parser they were handed rather
 * than one the runtime happened to have. A `ReferenceError` out of the middle of a read is what
 * this replaces: a caller could not tell it apart from a file that is damaged.
 */
describe("with no DOMParser global", () => {
  // Taken while the global is still there, the way a server takes one off jsdom
  const xmlParser = new DOMParser();

  beforeEach(() => {
    vi.stubGlobal("DOMParser", undefined);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("refuses to open a file with no-xml-parser rather than a ReferenceError", () => {
    expect(importErrorCode(() => importDocx(readFixture(FIXTURE)))).toBe(
      "no-xml-parser"
    );
  });

  it("opens the fixture through the xmlParser option instead", () => {
    const { doc, session } = importDocx(readFixture(FIXTURE), { xmlParser });

    expect(doc.textContent).not.toBe("");
    expect(documentPartPath(session)).toBe("word/document.xml");
  });

  /**
   * Writing reads the exported body back, and that read is the one wrapped in a refusal about the
   * document. The runtime's own refusal is settled as the call comes in, ahead of the wrapper, so
   * it comes through as itself even for a package that parses nothing until then
   */
  it("refuses to write a file with no-xml-parser rather than an export code", () => {
    const bare = makeDocx(`<w:p>${LETTER_SECT_PR}</w:p>`);
    const { doc, session } = importDocx(bare, { xmlParser });

    expect(importErrorCode(() => exportDocx(doc, session))).toBe(
      "no-xml-parser"
    );
  });

  /**
   * The parser is settled before the bytes are looked at, so a runtime that cannot read any file
   * says so rather than passing judgement on the one it was handed
   */
  it("refuses bytes that are no docx for the parser rather than for the bytes", () => {
    expect(importErrorCode(() => importDocx(new Uint8Array([1, 2, 3])))).toBe(
      "no-xml-parser"
    );
    expect(
      importErrorCode(() =>
        importDocx(new Uint8Array([1, 2, 3]), { xmlParser })
      )
    ).toBe("not-a-docx");
  });

  /**
   * A parser that gives out is a runtime problem, and letting its own exception through would put
   * the caller back where a bare `ReferenceError` left them: unable to place what went wrong.
   * There is nothing to read past a parser that refuses to read, so it reads as a file that did
   * not parse
   */
  it("refuses a file its parser threw on rather than letting the throw out", () => {
    const throwing: XmlParser = {
      parseFromString: () => {
        throw new TypeError("this parser gave up");
      },
    };

    expect(
      importErrorCode(() =>
        importDocx(readFixture(FIXTURE), { xmlParser: throwing })
      )
    ).toBe("malformed-xml");
  });

  /**
   * A session hands its numbering out through a reader of its own, and the parser it was opened
   * through is long out of scope by the time a consumer asks for it
   */
  it("reads an opened document's numbering through the same option", () => {
    const { session } = importDocx(readFixture(FIXTURE), { xmlParser });

    expect(
      documentNumbering(session, { xmlParser }).lists.size
    ).toBeGreaterThan(0);
    expect(importErrorCode(() => documentNumbering(session))).toBe(
      "no-xml-parser"
    );
  });

  it("reads a numbering part handed over on its own through the same option", () => {
    expect(parseNumbering(ONE_LIST_NUMBERING, { xmlParser }).lists.size).toBe(
      1
    );
    expect(importErrorCode(() => parseNumbering(ONE_LIST_NUMBERING))).toBe(
      "no-xml-parser"
    );
  });

  // A document with no numbering part defines no lists, and there is nothing there to read
  it("answers a document carrying no numbering with no parser at all", () => {
    const bare = makeDocx(`<w:p>${LETTER_SECT_PR}</w:p>`);
    const { session } = importDocx(bare, { xmlParser });

    expect(documentNumbering(session).lists.size).toBe(0);
    expect(parseNumbering(null).lists.size).toBe(0);
  });

  it("writes the file back out with the same option", () => {
    const { doc, session } = importDocx(readFixture(FIXTURE), { xmlParser });

    const out = exportDocx(doc, session, { xmlParser });

    expect(importDocx(out, { xmlParser }).doc.textContent).toBe(
      doc.textContent
    );
  });

  // The verifier opens two files through `importDocx`, and neither call is handed the option;
  // the scope the verifier put the parser in is what both of them read through
  it("takes the same option for a verdict and keeps it through both imports", () => {
    const bytes = readFixture(FIXTURE);

    expect(onlyCommentsChangedBy(bytes, bytes, "me", { xmlParser })).toEqual({
      ok: true,
    });
    expect(
      importErrorCode(() => onlyCommentsChangedBy(bytes, bytes, "me"))
    ).toBe("no-xml-parser");
  });
});
