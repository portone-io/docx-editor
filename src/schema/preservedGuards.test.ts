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
