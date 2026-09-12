// @vitest-environment jsdom
import { history, undo, undoDepth } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Plugin } from "prosemirror-state";
import { describe, expect, it, vi } from "vitest";
import { makeNotesDocx } from "../../__testing__/docx";
import { rangeOfText, runCommand } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { noteLabelsIn } from "../../docx/notes/numbering";
import { withSectionBreak } from "../../docx/sections";
import { editorStateForSession } from "../createEditor";
import { editorDocument, editorDocumentOf } from "../editorDocument";
import { noteNumbering } from "./noteNumbering";

const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const footnote = (id: string) =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

const THREE_FOOTNOTES =
  `<w:p>${footnote("2")}${text("between")}${footnote("3")}</w:p>` +
  `<w:p>${footnote("4")}</w:p>`;

/**
 * A state holding the snapshot, the history, and the numbering plugin alone.
 *
 * A full editing state keeps every note reference where the file put it, so deleting one is
 * refused there before the plugin could answer it.
 */
function numberedState(
  body: string,
  plugin: Plugin = noteNumbering()
): EditorState {
  const { doc, session } = importDocx(makeNotesDocx(body));
  return EditorState.create({
    doc,
    plugins: [
      editorDocument(editorDocumentOf(session, doc)),
      history(),
      plugin,
    ],
  });
}

/** Each reference as its id and the label it is drawn with, in document order */
function labels(doc: PMNode): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "noteReference") {
      found.push(`${node.attrs.id}:${node.attrs.label}`);
    }
    return true;
  });
  return found;
}

function referenceAt(doc: PMNode, id: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (
      found < 0 &&
      node.type.name === "noteReference" &&
      node.attrs.id === id
    ) {
      found = pos;
    }
    return found < 0;
  });
  if (found < 0) throw new Error(`no reference to note ${id}`);
  return found;
}

function withoutReference(state: EditorState, id: string): EditorState {
  const at = referenceAt(state.doc, id);
  return state.apply(state.tr.delete(at, at + 1));
}

describe("noteNumbering", () => {
  it("relabels every later footnote when a footnote reference is deleted", () => {
    const state = numberedState(THREE_FOOTNOTES);

    expect(labels(state.doc)).toEqual(["2:1", "3:2", "4:3"]);
    expect(labels(withoutReference(state, "2").doc)).toEqual(["3:1", "4:2"]);
  });

  it("relabels a reference inside a locked content control", () => {
    const eachSection =
      '<w:sectPr><w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr></w:sectPr>';
    const state = editorStateForSession(
      importDocx(
        makeNotesDocx(
          `<w:p>${footnote("2")}${text("open")}</w:p>` +
            '<w:p><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' +
            `${text("locked")}${footnote("3")}` +
            "</w:sdtContent></w:sdt></w:p>" +
            eachSection
        )
      )
    );
    const locked = rangeOfText(state.doc, "locked");
    const first = state.doc.child(0);
    const pPr: unknown = first.attrs.pPr;

    expect(
      state.apply(state.tr.insertText("x", locked.from + 2)).doc.eq(state.doc)
    ).toBe(true);
    expect(labels(state.doc)).toEqual(["2:1", "3:2"]);

    const broken = state.apply(
      state.tr.setNodeMarkup(0, undefined, {
        ...first.attrs,
        pPr: withSectionBreak(
          typeof pPr === "string" ? pPr : null,
          "<w:sectPr/>"
        ),
      })
    );

    expect(labels(broken.doc)).toEqual(["2:1", "3:1"]);
  });

  it("takes a relabel back with the edit that caused it in one undo", () => {
    const deleted = withoutReference(numberedState(THREE_FOOTNOTES), "2");

    expect(undoDepth(deleted)).toBe(1);

    const undone = runCommand(deleted, undo);

    expect(labels(undone.doc)).toEqual(["2:1", "3:2", "4:3"]);
    expect(undoDepth(undone)).toBe(0);
  });

  it("does not work labels out for an edit that reaches no reference or section break", () => {
    const labelsIn = vi.fn(noteLabelsIn);
    const state = numberedState(THREE_FOOTNOTES, noteNumbering(labelsIn));
    const typed = state.apply(
      state.tr.insertText("typed ", rangeOfText(state.doc, "between").from)
    );
    const reference = referenceAt(typed.doc, "3");
    const rewritten = typed.apply(
      typed.tr.setNodeAttribute(reference, "referenceXml", null)
    );

    expect(labelsIn).not.toHaveBeenCalled();
    expect(labels(rewritten.doc)).toEqual(["2:1", "3:2", "4:3"]);

    withoutReference(rewritten, "2");

    expect(labelsIn).toHaveBeenCalledTimes(1);
  });
});
