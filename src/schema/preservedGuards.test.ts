// @vitest-environment jsdom
import { joinBackward, joinForward } from "prosemirror-commands";
import { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { describe, expect, it, vi } from "vitest";
import { makeDocx, makeNotesDocx, NOTE_BODY } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "../editor/createEditor";
import { deleteColumn, deleteRow, deleteTable } from "../table";
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

const TABLE_GRID = '<w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>';
const TABLE_CELL = `<w:tc><w:p>${runXml("cell")}</w:p></w:tc>`;
const BOOKMARK_START = '<w:bookmarkStart w:id="9" w:name="b"/>';
const BOOKMARK_END = '<w:bookmarkEnd w:id="9"/>';

function tableMarkerState(
  position: "table" | "rowStart" | "cell" | "rowEnd",
  marker = BOOKMARK_START
): EditorState {
  const body =
    `<w:tbl>${TABLE_GRID}${position === "table" ? marker : ""}` +
    `<w:tr>${position === "rowStart" ? marker : ""}${TABLE_CELL}${position === "cell" ? marker : ""}</w:tr>` +
    `${position === "rowEnd" ? marker : ""}</w:tbl><w:p>${BOOKMARK_END}${runXml("after")}</w:p>`;
  const state = bookmarked(body);
  const cell = nodeRange(state.doc, "tableCell");
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, cell.from + 2))
  );
}

describe("markers stored around table content", () => {
  it.each([
    ["table", "table", "leadingXml"],
    ["rowStart", "tableRow", "leadingXml"],
    ["cell", "tableCell", "trailingXml"],
    ["rowEnd", "tableRow", "trailingXml"],
  ] as const)(
    "guards the %s carrier while allowing typing inside it",
    (position, type, attr) => {
      const state = tableMarkerState(position);
      const carrier = nodeRange(state.doc, type);
      expect(deleteTable(state)).toBe(false);
      expect(
        transactionAllowed(
          state.tr.setNodeAttribute(carrier.from, attr, null),
          state
        )
      ).toBe(false);
      expect(editShut(state, { kind: "replace", ...carrier })).toBe(true);
      const at = state.selection.from;
      expect(editShut(state, { kind: "replace", from: at, to: at + 1 })).toBe(
        false
      );
      const walked = vi.spyOn(PMNode.prototype, "descendants");
      try {
        expect(transactionAllowed(state.tr.insertText("x", at), state)).toBe(
          true
        );
        expect(walked).not.toHaveBeenCalled();
      } finally {
        walked.mockRestore();
      }
    }
  );

  it("lets a row carrying only a proofing trace be deleted", () => {
    let state = bookmarked(
      `<w:tbl>${TABLE_GRID}<w:tr><w:proofErr w:type="spellStart"/>${TABLE_CELL}</w:tr><w:tr>${TABLE_CELL}</w:tr></w:tbl><w:p/>`
    );
    const cell = nodeRange(state.doc, "tableCell");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, cell.from + 2))
    );
    expect(deleteRow(state)).toBe(true);
    const row = nodeRange(state.doc, "tableRow");
    expect(editShut(state, { kind: "replace", ...row })).toBe(false);
  });

  it("does not mistake an explicitly foreign marker name for a Word marker", () => {
    const state = tableMarkerState("rowStart");
    const row = nodeRange(state.doc, "tableRow");
    const foreign = state.doc.nodeAt(row.from)?.type.create(
      {
        ...state.doc.nodeAt(row.from)?.attrs,
        leadingXml: '<other:bookmarkStart xmlns:other="urn:other"/>',
      },
      state.doc.nodeAt(row.from)?.content
    );
    if (!foreign) throw new Error("no row");
    const doc = state.tr.replaceWith(row.from, row.to, foreign).doc;
    const foreignState = createEditorState(doc);
    expect(
      transactionAllowed(foreignState.tr.delete(row.from, row.to), foreignState)
    ).toBe(true);
  });

  it("refuses planting another marker on an empty carrier", () => {
    const state = tableMarkerState("cell");
    const row = nodeRange(state.doc, "tableRow");
    expect(
      transactionAllowed(
        state.tr.setNodeAttribute(row.from, "leadingXml", BOOKMARK_START),
        state
      )
    ).toBe(false);
  });

  it("keeps the order between a trailing marker and an inline marker", () => {
    const state = bookmarked(
      `<w:tbl>${TABLE_GRID}<w:tr><w:tc><w:p>${BOOKMARK_START}${runXml("cell")}</w:p></w:tc>${BOOKMARK_END}</w:tr></w:tbl><w:p>${runXml("after")}</w:p>`
    );
    const start = nodeRange(state.doc, "rawInline");
    const marker = state.doc.nodeAt(start.from);
    if (!marker) throw new Error("no marker");
    const tr = state.tr.delete(start.from, start.to);
    tr.insert(tr.doc.child(0).nodeSize + 1, marker);
    expect(transactionAllowed(tr, state)).toBe(false);
  });

  it("allows moving a table whole with its marker pair", () => {
    const state = bookmarked(
      `<w:p>${runXml("before")}</w:p><w:tbl>${TABLE_GRID}<w:tr><w:tc><w:p>${BOOKMARK_START}${runXml("cell")}</w:p></w:tc>${BOOKMARK_END}</w:tr></w:tbl>`
    );
    const table = nodeRange(state.doc, "table");
    const moved = state.doc.slice(table.from, table.to);
    const tr = state.tr.delete(table.from, table.to).replace(0, 0, moved);
    expect(transactionAllowed(tr, state)).toBe(true);
    expect(state.apply(tr).doc.eq(tr.doc)).toBe(true);
    expect(transactionAllowed(state.tr.replace(0, 0, moved), state)).toBe(
      false
    );
  });
});

describe("paragraph joins beside body bookmark placeholders", () => {
  it.each(["backward", "forward"] as const)(
    "keeps the body marker on a %s join",
    (direction) => {
      let state = bookmarked(
        `<w:p>${runXml("before")}</w:p>${BOOKMARK_START}<w:p>${runXml("inside")}</w:p>${BOOKMARK_END}<w:p>${runXml("after")}</w:p>`
      );
      const marker = nodeRange(state.doc, "rawBlock");
      const at = direction === "backward" ? marker.to + 1 : marker.from - 1;
      state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, at))
      );
      const before = state.doc;
      const command = direction === "backward" ? joinBackward : joinForward;
      command(state, (tr) => {
        state = state.apply(tr);
      });
      expect(state.doc.eq(before)).toBe(true);
      expect(countOf(state.doc, "rawBlock")).toBe(2);
    }
  );
});

describe("a guard over the markers a document was opened with", () => {
  it.each([deleteRow, deleteColumn])(
    "refuses removing the cell carrying a table bookmark endpoint",
    (command) => {
      let state = bookmarked(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
          "<w:tr><w:tc><w:p>" +
          runXml("a") +
          '</w:p></w:tc><w:bookmarkStart w:id="9" w:name="b"/>' +
          "<w:tc><w:p>" +
          runXml("b") +
          "</w:p></w:tc></w:tr>" +
          "<w:tr><w:tc><w:p>" +
          runXml("c") +
          "</w:p></w:tc><w:tc><w:p>" +
          runXml("d") +
          "</w:p></w:tc></w:tr>" +
          '</w:tbl><w:p><w:bookmarkEnd w:id="9"/>' +
          runXml("end") +
          "</w:p>"
      );
      const cell = nodeRange(state.doc, "tableCell");
      state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, cell.from + 2))
      );
      expect(command(state)).toBe(false);
      const dispatch = vi.fn();
      expect(command(state, dispatch)).toBe(false);
      expect(dispatch).not.toHaveBeenCalled();
    }
  );

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

  it("allows moving a complete field to every position in the surrounding text", () => {
    const state = bookmarked(
      `<w:p>${runXml("alpha")}<w:r>` +
        '<w:fldChar w:fldCharType="begin"/>' +
        "<w:instrText> PAGE </w:instrText>" +
        '<w:fldChar w:fldCharType="end"/></w:r>' +
        `${runXml("omega")}</w:p>`
    );
    const { from } = nodeRange(state.doc, "rawRunContent");
    const field = state.doc.slice(from, from + 3);
    const removed = state.tr.delete(from, from + 3);
    expect(transactionAllowed(removed, state)).toBe(false);
    for (
      let destination = 1;
      destination < removed.doc.child(0).nodeSize;
      destination++
    ) {
      const tr = state.tr.delete(from, from + 3);
      tr.replace(destination, destination, field);
      expect(transactionAllowed(tr, state), `destination ${destination}`).toBe(
        true
      );
      expect(state.apply(tr).doc.eq(tr.doc)).toBe(true);
    }
  });

  it("refuses reordering field pieces even when none is lost", () => {
    const state = bookmarked(
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/>' +
        "<w:instrText> PAGE </w:instrText>" +
        '<w:fldChar w:fldCharType="end"/></w:r></w:p>'
    );
    const { from, to } = nodeRange(state.doc, "rawRunContent");
    const begin = state.doc.slice(from, to);
    const tr = state.tr.delete(from, to);
    tr.replace(3, 3, begin);
    expect(transactionAllowed(tr, state)).toBe(false);
    expect(state.apply(tr).doc.eq(state.doc)).toBe(true);
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
