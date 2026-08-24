// @vitest-environment jsdom
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  exportErrorCode,
  fixtureNames,
  makeDocx,
  makeNumberedDocx,
  ONE_LIST_NUMBERING,
  readFixture,
} from "../__testing__/docx";
import {
  toggleBulletList,
  toggleNumberedList,
} from "../editor/commands/listCommands";
import { createEditorState } from "../editor/createEditor";
import { paragraphMarkers } from "../editor/plugins/numberingDecorations";
import { toParagraphFormat } from "../model/format";
import { parseNumbering } from "../numbering/parseNumbering";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import type { SessionStore } from "./session";

const NUMBERING_PART = "word/numbering.xml";

/** The first paragraph that is not a list and has text, plus a position inside it */
function plainParagraph(doc: PMNode): { index: number; pos: number } {
  let spot: { index: number; pos: number } | null = null;
  doc.forEach((block, offset, index) => {
    if (spot !== null || block.type.name !== "paragraph") return;
    if (block.textContent.length === 0) return;
    if (toParagraphFormat(block.attrs.format)?.numbering) return;
    spot = { index, pos: offset + 1 };
  });
  if (spot === null) throw new Error("no paragraph outside a list");
  return spot;
}

function withCaretAt(state: EditorState, at: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at))
  );
}

function open(name: string) {
  const bytes = readFixture(name);
  const { doc, session } = importDocx(bytes);
  const numbering = parseNumbering(session.numberingXml);
  return { bytes, doc, session, state: createEditorState(doc, { numbering }) };
}

function partsOf(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

/** Exports a document in which a new list has been started */
function exportWithNewList(name: string): {
  bytes: Uint8Array;
  session: SessionStore;
  /** Which block in the body the paragraph that became a list is */
  index: number;
  out: Uint8Array;
} {
  const { bytes, doc, session, state } = open(name);
  const spot = plainParagraph(doc);
  const at = withCaretAt(state, spot.pos);
  let listed = at;
  expect(
    toggleNumberedList(at, (tr) => {
      listed = at.apply(tr);
    })
  ).toBe(true);
  return {
    bytes,
    session,
    index: spot.index,
    out: exportDocx(listed.doc, session),
  };
}

describe("editing that starts a new list", () => {
  it.each(fixtureNames)(
    "%s: numbering.xml keeps the original text intact",
    (name) => {
      const { bytes, out } = exportWithNewList(name);
      const before = decode(partsOf(bytes)[NUMBERING_PART]);
      const after = decode(partsOf(out)[NUMBERING_PART]);

      const firstNumAt = before.indexOf("<w:num ");
      const closeAt = before.lastIndexOf("</w:numbering>");
      expect(firstNumAt).toBeGreaterThan(0);

      // The definition is spliced in before the first number and the number at the very end, so the original text survives as three intact chunks
      expect(after.startsWith(before.slice(0, firstNumAt))).toBe(true);
      expect(after).toContain(before.slice(firstNumAt, closeAt));
      expect(after.endsWith(before.slice(closeAt))).toBe(true);
      expect(after.length).toBeGreaterThan(before.length);
    }
  );

  it.each(fixtureNames)(
    "%s: the body regenerates only that paragraph",
    (name) => {
      const { out, session, index } = exportWithNewList(name);
      const documentXml = decode(partsOf(out)[session.mainPartPath]);

      const head =
        session.documentPrefix +
        session.blocks
          .slice(0, index)
          .map((block) => block.xml)
          .join("");
      const tail =
        session.blocks
          .slice(index + 1)
          .map((block) => block.xml)
          .join("") + session.documentSuffix;
      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);
      expect(
        documentXml.slice(head.length, documentXml.length - tail.length)
      ).toContain("<w:numPr>");
    }
  );

  it.each(fixtureNames)(
    "%s: only the body and numbering.xml change",
    (name) => {
      const { bytes, out, session } = exportWithNewList(name);
      const original = partsOf(bytes);
      const exported = partsOf(out);

      for (const key of Object.keys(original)) {
        if (key === session.mainPartPath || key === NUMBERING_PART) continue;
        expect(bytesEqual(exported[key], original[key])).toBe(true);
      }
    }
  );

  it.each(fixtureNames)(
    "%s: reopening it computes the numbers of the new list",
    (name) => {
      const { out } = exportWithNewList(name);
      const again = importDocx(out);
      const numbering = parseNumbering(again.session.numberingXml);
      const markers = paragraphMarkers(again.doc, numbering);

      // The new list is a single paragraph, so the first number shows up as is
      const newList = markers.filter((marker) => marker.text === "1.");
      expect(newList.length).toBeGreaterThan(0);
    }
  );
});

function listParagraph(numId: number, text: string): string {
  return (
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>` +
    `<w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

describe("a list that had no definition from the moment it was opened", () => {
  const bytes = makeNumberedDocx(
    listParagraph(9, "list with no definition"),
    ONE_LIST_NUMBERING
  );

  it("exporting without editing does not disturb numbering.xml", () => {
    const { doc, session } = importDocx(bytes);
    const exported = partsOf(exportDocx(doc, session));
    expect(
      bytesEqual(exported[NUMBERING_PART], partsOf(bytes)[NUMBERING_PART])
    ).toBe(true);
  });

  it("editing the text does not create a new definition", () => {
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("edit", 2));
    const exported = partsOf(exportDocx(edited.doc, session));

    expect(
      bytesEqual(exported[NUMBERING_PART], partsOf(bytes)[NUMBERING_PART])
    ).toBe(true);
  });

  it("starting a new list in that document adds only the new definition", () => {
    const withPlain = makeNumberedDocx(
      listParagraph(9, "list with no definition") +
        '<w:p><w:r><w:t xml:space="preserve">plain paragraph</w:t></w:r></w:p>',
      ONE_LIST_NUMBERING
    );
    const { doc, session } = importDocx(withPlain);
    const state = createEditorState(doc, {
      numbering: parseNumbering(session.numberingXml),
    });
    const at = withCaretAt(state, plainParagraph(doc).pos);
    let listed = at;
    expect(
      toggleNumberedList(at, (tr) => {
        listed = at.apply(tr);
      })
    ).toBe(true);

    const numbering = decode(
      partsOf(exportDocx(listed.doc, session))[NUMBERING_PART]
    );
    const added = parseNumbering(numbering).lists;
    // Only the already-defined 1 and the newly started list are there. 9 still has no definition
    expect(added.has(9)).toBe(false);
    expect(added.size).toBe(2);
  });
});

describe("starting two lists in one document", () => {
  it("the numbered list and the bullet list each get their own definition", () => {
    const bytes = makeNumberedDocx(
      '<w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">second</w:t></w:r></w:p>',
      ONE_LIST_NUMBERING
    );
    const { doc, session } = importDocx(bytes);
    const numbering = parseNumbering(session.numberingXml);
    const state = createEditorState(doc, { numbering });

    const first = withCaretAt(state, 1);
    let numbered = first;
    expect(
      toggleNumberedList(first, (tr) => {
        numbered = first.apply(tr);
      })
    ).toBe(true);

    const second = withCaretAt(numbered, numbered.doc.child(0).nodeSize + 1);
    let bulleted = second;
    expect(
      toggleBulletList(second, (tr) => {
        bulleted = second.apply(tr);
      })
    ).toBe(true);

    const out = exportDocx(bulleted.doc, session);
    const lists = parseNumbering(decode(partsOf(out)[NUMBERING_PART])).lists;
    expect(lists.size).toBe(3);

    const again = importDocx(out);
    const markers = paragraphMarkers(
      again.doc,
      parseNumbering(again.session.numberingXml)
    );
    expect(markers.map((marker) => marker.text)).toEqual(["1.", "●"]);
  });
});

describe("a document without numbering.xml", () => {
  it("stops when it tries to add a new list", () => {
    const { doc, session } = importDocx(
      makeDocx('<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>')
    );
    const state = createEditorState(doc);
    const at = withCaretAt(state, 1);
    let listed = at;
    expect(
      toggleNumberedList(at, (tr) => {
        listed = at.apply(tr);
      })
    ).toBe(true);

    expect(exportErrorCode(() => exportDocx(listed.doc, session))).toBe(
      "missing-numbering-part"
    );
  });
});
