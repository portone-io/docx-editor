// @vitest-environment jsdom
import { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it, vi } from "vitest";
import { makeDocx, makeNotesDocx, NOTE_BODY } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "../editor/createEditor";
import { editShut, transactionAllowed } from "./guards";

const runXml = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** One paragraph reading "mkn" with a bookmark range anchored around the "k" */
const BOOKMARK_P =
  "<w:p>" +
  runXml("m") +
  '<w:bookmarkStart w:id="9" w:name="b"/>' +
  runXml("k") +
  '<w:bookmarkEnd w:id="9"/>' +
  runXml("n") +
  "</w:p>";

function bookmarked(body = BOOKMARK_P): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc);
}

function noted(): EditorState {
  return createEditorState(importDocx(makeNotesDocx(NOTE_BODY)).doc);
}

/** The stretch the first node of this type covers */
function nodeRange(
  doc: PMNode,
  typeName: string
): { from: number; to: number } {
  const found: { from: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (found.length === 0 && node.type.name === typeName) {
      found.push({ from: pos, to: pos + node.nodeSize });
    }
    return found.length === 0;
  });
  const first = found[0];
  if (first === undefined) throw new Error(`no ${typeName} in the document`);
  return first;
}

/** How many nodes of this type the document holds */
function countOf(doc: PMNode, typeName: string): number {
  let seen = 0;
  doc.descendants((node) => {
    if (node.type.name === typeName) seen += 1;
    return true;
  });
  return seen;
}

describe("a guard over the markers a document was opened with", () => {
  it("refuses a step that takes a bookmark marker away", () => {
    const state = bookmarked();
    const marker = nodeRange(state.doc, "rawInline");
    expect(
      transactionAllowed(state.tr.delete(marker.from, marker.to), state)
    ).toBe(false);
  });

  it("refuses a transaction that drops one half of a bookmark", () => {
    const state = bookmarked();
    const marker = nodeRange(state.doc, "rawInline");

    expect(
      transactionAllowed(state.tr.delete(marker.from, marker.to), state)
    ).toBe(false);
  });

  it("refuses a transaction that drops a field character", () => {
    const state = bookmarked(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/>` +
        '<w:instrText xml:space="preserve"> PAGE </w:instrText>' +
        '<w:fldChar w:fldCharType="end"/></w:r>' +
        `${runXml("n")}</w:p>`
    );
    const piece = nodeRange(state.doc, "rawRunContent");

    expect(
      transactionAllowed(state.tr.delete(piece.from, piece.to), state)
    ).toBe(false);
    expect(editShut(state, { kind: "replace", ...piece })).toBe(true);
  });

  it("lets a transaction delete a revision chip whole", () => {
    const state = bookmarked(
      `<w:p>${runXml("m")}<w:ins w:id="1" w:author="Reviewer A" ` +
        `w:date="2026-01-01T00:00:00Z">${runXml("k")}</w:ins>` +
        `${runXml("n")}</w:p>`
    );
    const chip = nodeRange(state.doc, "rawInline");

    expect(transactionAllowed(state.tr.delete(chip.from, chip.to), state)).toBe(
      true
    );
    expect(editShut(state, { kind: "replace", ...chip })).toBe(false);
  });

  it("lets a transaction delete a proofErr", () => {
    const state = bookmarked(
      `<w:p><w:proofErr w:type="spellStart"/>${runXml("m")}</w:p>`
    );
    const mark = nodeRange(state.doc, "rawInline");

    expect(transactionAllowed(state.tr.delete(mark.from, mark.to), state)).toBe(
      true
    );
  });

  it("refuses a step that plants a second note reference", () => {
    const state = noted();
    const { from } = nodeRange(state.doc, "noteReference");
    const reference = state.doc.nodeAt(from);
    if (reference === null) throw new Error("no note reference");

    expect(transactionAllowed(state.tr.insert(1, reference), state)).toBe(
      false
    );
  });

  /**
   * The whole list of markers is compared by walking the document, and the point of the guard is
   * that ordinary typing never reaches that walk.
   */
  it("lets a step that rewrites text beside a marker through without walking the document", () => {
    const state = bookmarked();
    const marker = nodeRange(state.doc, "rawInline");
    const walked = vi.spyOn(PMNode.prototype, "descendants");
    try {
      expect(transactionAllowed(state.tr.insertText("x", 1), state)).toBe(true);
      expect(walked).not.toHaveBeenCalled();

      // The very same guard does walk it once a step has reached the marker
      expect(
        transactionAllowed(state.tr.delete(marker.from, marker.to), state)
      ).toBe(false);
      expect(walked).toHaveBeenCalled();
    } finally {
      walked.mockRestore();
    }
  });

  /**
   * What a drag writes: one step takes the paragraph out and one puts it back, and the first of
   * them on its own leaves the document without the markers the second returns.
   */
  it("keeps a marker moved whole in the same order", () => {
    const state = bookmarked(`<w:p>${runXml("alpha")}</w:p>${BOOKMARK_P}`);
    const from = state.doc.child(0).nodeSize;
    const to = from + state.doc.child(1).nodeSize;
    const moved = state.doc.slice(from, to);

    const tr = state.tr.delete(from, to);
    tr.replace(0, 0, moved);

    expect(tr.steps).toHaveLength(2);
    expect(transactionAllowed(tr, state)).toBe(true);
    expect(tr.doc.child(0).textContent).toBe("mkn");
  });

  it("refuses a marker moved out of the order it stood in", () => {
    const state = bookmarked();
    const start = nodeRange(state.doc, "rawInline");
    const marker = state.doc.nodeAt(start.from);
    if (marker === null) throw new Error("no bookmark marker");

    // The opening marker is taken from where it stood and put back after the closing one
    const tr = state.tr.delete(start.from, start.to);
    tr.insert(tr.doc.child(0).nodeSize - 1, marker);

    expect(countOf(tr.doc, "rawInline")).toBe(countOf(state.doc, "rawInline"));
    expect(transactionAllowed(tr, state)).toBe(false);
  });

  it("shuts a replace intent over a marker and leaves an insert beside it open", () => {
    const withBookmark = bookmarked();
    const marker = nodeRange(withBookmark.doc, "rawInline");
    expect(editShut(withBookmark, { kind: "replace", ...marker })).toBe(true);
    expect(editShut(withBookmark, { kind: "insert", at: marker.from })).toBe(
      false
    );
    expect(editShut(withBookmark, { kind: "insert", at: marker.to })).toBe(
      false
    );
    // A mark laid over a marker leaves it standing, so it is nobody's business here
    expect(editShut(withBookmark, { kind: "mark", ...marker })).toBe(false);

    const withNote = noted();
    const reference = nodeRange(withNote.doc, "noteReference");
    expect(editShut(withNote, { kind: "replace", ...reference })).toBe(true);
    expect(editShut(withNote, { kind: "insert", at: reference.to })).toBe(
      false
    );
  });
});

describe("a guard over the sections a document was opened with", () => {
  const SECT_PR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

  /** Two paragraphs, the second of which ends a section */
  function sectioned(): EditorState {
    return createEditorState(
      importDocx(
        makeDocx(
          `<w:p>${runXml("one")}</w:p>` +
            `<w:p><w:pPr>${SECT_PR}</w:pPr>${runXml("two")}</w:p>`
        )
      ).doc
    );
  }

  /** Where the paragraph laying down the section break starts */
  function sectionParagraphAt(doc: PMNode): number {
    let found = -1;
    doc.forEach((node, offset) => {
      if (found < 0 && typeof node.attrs.pPr === "string") found = offset;
    });
    if (found < 0) throw new Error("no paragraph carrying a section break");
    return found;
  }

  it("refuses a transaction that joins away a paragraph carrying a section break", () => {
    const state = sectioned();
    const at = sectionParagraphAt(state.doc);

    // The boundary between the two paragraphs goes, which leaves the first one's attrs standing
    const tr = state.tr.delete(at - 1, at + 1);

    expect(tr.doc.childCount).toBe(1);
    expect(transactionAllowed(tr, state)).toBe(false);
  });

  it("lets a section paragraph be split", () => {
    const state = sectioned();
    const at = sectionParagraphAt(state.doc);
    const paragraph = state.doc.child(state.doc.childCount - 1);

    // What the keymap builds: the break goes to the half that now ends the section
    const tr = state.tr.split(at + 2, 1, [
      { type: paragraph.type, attrs: paragraph.attrs },
    ]);
    tr.setNodeMarkup(at, null, { ...paragraph.attrs, pPr: null });

    expect(tr.doc.childCount).toBe(3);
    expect(transactionAllowed(tr, state)).toBe(true);
  });

  it("leaves a document without section paragraphs alone", () => {
    const state = createEditorState(
      importDocx(makeDocx(`<w:p>${runXml("one")}</w:p>`)).doc
    );

    expect(transactionAllowed(state.tr.delete(1, 4), state)).toBe(true);
  });

  it("shuts a replace intent over a section paragraph and leaves an insert inside it open", () => {
    const state = sectioned();
    const at = sectionParagraphAt(state.doc);
    const paragraph = state.doc.child(state.doc.childCount - 1);
    const range = { from: at, to: at + paragraph.nodeSize };

    expect(editShut(state, { kind: "replace", ...range })).toBe(true);
    // Typing inside such a paragraph takes nothing away, so the guard stands aside
    expect(editShut(state, { kind: "insert", at: at + 1 })).toBe(false);
    expect(editShut(state, { kind: "mark", ...range })).toBe(false);
  });

  it("allows replacing text inside a section paragraph", () => {
    const state = sectioned();
    const at = sectionParagraphAt(state.doc);

    expect(editShut(state, { kind: "replace", from: at + 1, to: at + 4 })).toBe(
      false
    );
    expect(
      transactionAllowed(state.tr.insertText("new", at + 1, at + 4), state)
    ).toBe(true);
  });

  it("does not read the section break a tracked change kept", () => {
    const state = createEditorState(
      importDocx(
        makeDocx(
          `<w:p>${runXml("one")}</w:p>` +
            "<w:p><w:pPr>" +
            '<w:pPrChange w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">' +
            `<w:pPr>${SECT_PR}</w:pPr></w:pPrChange>` +
            `</w:pPr>${runXml("two")}</w:p>`
        )
      ).doc
    );

    // The break belongs to the properties the paragraph wore before the change, not to it
    expect(transactionAllowed(state.tr.delete(4, 6), state)).toBe(true);
  });
});
