// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  makeNotesDocx,
  NOTE_BODY,
} from "../__testing__/docx";
import { documentNotes } from "../editor/commands/noteQueries";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

function editFirstText(doc: PMNode): PMNode {
  const paragraph = doc.child(0);
  const inline: PMNode[] = [];
  let edited = false;
  paragraph.forEach((child) => {
    if (!edited && child.isText) {
      inline.push(docxSchema.text("Edited text", child.marks));
      edited = true;
    } else {
      inline.push(child);
    }
  });
  return docxSchema.nodes.doc.create(null, [
    paragraph.copy(Fragment.from(inline)),
  ]);
}

describe("footnotes and endnotes", () => {
  it("imports references without replacing their paragraph", () => {
    const { doc } = importDocx(makeNotesDocx());

    expect(doc.child(0).type.name).toBe("paragraph");
    expect(doc.child(0).children.map((node) => node.type.name)).toEqual([
      "text",
      "noteReference",
      "text",
      "noteReference",
    ]);
    expect(doc.textContent).toBe("Text and more");
  });

  it("reads regular note bodies and omits separator notes from the document query", () => {
    const state = createEditorState(importDocx(makeNotesDocx()).doc);

    expect(documentNotes(state)).toEqual([
      expect.objectContaining({
        kind: "footnote",
        id: "2",
        label: "1",
        text: "Footnote body\nSecond line",
      }),
      expect.objectContaining({
        kind: "endnote",
        id: "3",
        label: "1",
        text: "Endnote body",
      }),
    ]);
  });

  it("numbers notes by first reference rather than note-part order", () => {
    const bytes = makeNotesDocx(
      '<w:p><w:r><w:footnoteReference w:id="2"/></w:r>' +
        '<w:r><w:footnoteReference w:id="4"/></w:r></w:p>'
    );
    const parts = unzipSync(bytes);
    parts["word/footnotes.xml"] = new TextEncoder().encode(
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:footnote w:id="4"><w:p><w:r><w:t>Second body</w:t></w:r></w:p></w:footnote>' +
        '<w:footnote w:id="2"><w:p><w:r><w:t>First body</w:t></w:r></w:p></w:footnote>' +
        "</w:footnotes>"
    );

    expect(
      documentNotes(createEditorState(importDocx(zipSync(parts)).doc)).map(
        ({ id, label }) => ({ id, label })
      )
    ).toEqual([
      { id: "2", label: "1" },
      { id: "4", label: "2" },
    ]);
  });

  it("resolves a note part relationship containing parent path segments", () => {
    const parts = unzipSync(makeNotesDocx());
    parts["notes/footnotes.xml"] = parts["word/footnotes.xml"];
    delete parts["word/footnotes.xml"];
    parts["word/_rels/document.xml.rels"] = new TextEncoder().encode(
      decode(parts["word/_rels/document.xml.rels"]).replace(
        'Target="footnotes.xml"',
        'Target="../notes/footnotes.xml"'
      )
    );
    parts["[Content_Types].xml"] = new TextEncoder().encode(
      decode(parts["[Content_Types].xml"]).replace(
        "/word/footnotes.xml",
        "/notes/footnotes.xml"
      )
    );

    expect(
      documentNotes(createEditorState(importDocx(zipSync(parts)).doc))[0]?.text
    ).toContain("Footnote body");
  });

  it("refuses deleting or duplicating a display-only note reference", () => {
    const state = createEditorState(importDocx(makeNotesDocx()).doc);
    let position = -1;
    state.doc.descendants((node, pos) => {
      if (position < 0 && node.type.name === "noteReference") {
        position = pos;
      }
      return position < 0;
    });
    const reference = position < 0 ? null : state.doc.nodeAt(position);
    if (position < 0 || !reference) {
      throw new Error("no note reference in the test document");
    }

    expect(
      state
        .apply(state.tr.delete(position, position + reference.nodeSize))
        .doc.eq(state.doc)
    ).toBe(true);
    expect(
      state.apply(state.tr.insert(position, reference)).doc.eq(state.doc)
    ).toBe(true);
  });

  it("keeps note parts byte-identical after editing surrounding body text", () => {
    const bytes = makeNotesDocx();
    const before = unzipSync(bytes);
    const opened = importDocx(bytes);
    const output = exportDocx(editFirstText(opened.doc), opened.session);
    const after = unzipSync(output);

    expect(
      bytesEqual(after["word/footnotes.xml"], before["word/footnotes.xml"])
    ).toBe(true);
    expect(
      bytesEqual(after["word/endnotes.xml"], before["word/endnotes.xml"])
    ).toBe(true);
    const documentXml = decode(after["word/document.xml"]);
    expect(documentXml).toContain('<w:footnoteReference w:id="2"/>');
    expect(documentXml).toContain('<w:endnoteReference w:id="3"/>');
    expect(documentXml).toContain("Edited text");
  });

  it("round-trips all note package parts byte-identically when untouched", () => {
    const bytes = makeNotesDocx(NOTE_BODY);
    const opened = importDocx(bytes);
    const before = unzipSync(bytes);
    const after = unzipSync(exportDocx(opened.doc, opened.session));

    for (const path of [
      "word/document.xml",
      "word/footnotes.xml",
      "word/endnotes.xml",
      "word/_rels/document.xml.rels",
      "[Content_Types].xml",
    ]) {
      expect(bytesEqual(after[path], before[path]), path).toBe(true);
    }
  });
});
