// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { undoDepth } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { decode, makeNotesDocx } from "../../__testing__/docx";
import { rangeOfText, runCommand, select } from "../../__testing__/editing";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import { storyText } from "../../docx/story";
import {
  type NoteKind,
  sameStory,
  storiesOf,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { insertFootnote } from "../commands/footnoteCommands";
import { redo, undo } from "../commands/historyCommands";
import { documentNotes } from "../commands/noteQueries";
import { editorStateForSession } from "../createEditor";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const footnote = (id: string) =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
const endnote = (id: string) => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;

const entry = (id: string, words: string, paraId: string) =>
  `<w:footnote w:id="${id}"><w:p w14:paraId="${paraId}">` +
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
  `${text(` ${words}`)}</w:p></w:footnote>`;

/** Two footnotes the text calls, and one it does not */
const NOTES = [
  entry("1", "First note", "10000001"),
  entry("2", "Second note", "10000002"),
  entry("7", "Nobody calls this note", "10000007"),
];

const BODY =
  `<w:p>${text("One")}${footnote("1")}${text(" two")}${footnote("2")}</w:p>` +
  `<w:p>${text("Three")}</w:p>`;

/** A package whose Footnotes part holds a separator and these entries */
function footnotesDocx(body: string, entries: readonly string[]): Uint8Array {
  const parts = unzipSync(makeNotesDocx(body));
  parts["word/footnotes.xml"] = new TextEncoder().encode(
    `<w:footnotes xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      entries.join("") +
      "</w:footnotes>"
  );
  return zipSync(parts);
}

function opened(body = BODY): EditorState {
  return editorStateForSession(importDocx(footnotesDocx(body, NOTES)));
}

interface PlacedReference {
  readonly id: string;
  readonly pos: number;
  readonly node: PMNode;
}

function noteReferences(doc: PMNode, kind: NoteKind): PlacedReference[] {
  const found: PlacedReference[] = [];
  doc.descendants((node, pos) => {
    const id: unknown = node.attrs.id;
    if (
      node.type.name === "noteReference" &&
      node.attrs.kind === kind &&
      typeof id === "string"
    ) {
      found.push({ id, pos, node });
    }
    return true;
  });
  return found;
}

function footnoteReferences(doc: PMNode): PlacedReference[] {
  return noteReferences(doc, "footnote");
}

function referenceTo(state: EditorState, id: string): PlacedReference {
  const reference = footnoteReferences(state.doc).find(
    (candidate) => candidate.id === id
  );
  if (reference === undefined) throw new Error(`no reference to ${id}`);
  return reference;
}

function footnoteStory(state: EditorState, id: string): PMNode | null {
  return storyNodeOf(state.doc, storyKey("footnote", id));
}

function withoutReference(state: EditorState, id: string): EditorState {
  const { pos } = referenceTo(state, id);
  return state.apply(state.tr.delete(pos, pos + 1));
}

/** The reference copied to the start of the paragraph holding this text */
function copiedBefore(
  state: EditorState,
  id: string,
  needle: string
): EditorState {
  const { node } = referenceTo(state, id);
  return state.apply(
    state.tr.insert(rangeOfText(state.doc, needle).from, node)
  );
}

function footnoteKeys(state: EditorState): string[] {
  return Object.keys(storiesOf(state.doc))
    .filter((key) => key.startsWith("footnote:"))
    .sort();
}

/** The footnote references that name no story, which an edit may never leave behind */
function unnamed(state: EditorState): string[] {
  return footnoteReferences(state.doc)
    .filter(({ id }) => footnoteStory(state, id) === null)
    .map(({ id }) => id);
}

describe("noteLifecycle", () => {
  it("deletes a footnote's story in the transaction that deletes its last reference", () => {
    const state = opened();
    const { pos } = referenceTo(state, "2");
    const { state: after, transactions } = state.applyTransaction(
      state.tr.delete(pos, pos + 1)
    );

    expect(footnoteStory(after, "2")).toBeNull();
    expect(footnoteStory(after, "1")).not.toBeNull();
    expect(storyNodeOf(after.doc, storyKey("endnote", "3"))).not.toBeNull();
    expect(transactions).toHaveLength(2);
    expect(undoDepth(after)).toBe(1);
  });

  it("deletes an endnote's story in the transaction that deletes its last reference", () => {
    const state = opened(`${BODY}<w:p>${text("Four")}${endnote("3")}</w:p>`);
    const key = storyKey("endnote", "3");
    const reference = noteReferences(state.doc, "endnote")[0];
    if (reference === undefined) throw new Error("no endnote reference");
    expect(storyNodeOf(state.doc, key)).not.toBeNull();

    const { state: after, transactions } = state.applyTransaction(
      state.tr.delete(reference.pos, reference.pos + 1)
    );

    expect(storyNodeOf(after.doc, key)).toBeNull();
    expect(footnoteStory(after, "2")).not.toBeNull();
    expect(transactions).toHaveLength(2);
    expect(undoDepth(after)).toBe(1);
  });

  it("keeps the story while another reference to the same id remains", () => {
    const state = opened(
      `<w:p>${text("One")}${footnote("2")}${text(" two")}${footnote("2")}</w:p>`
    );
    const once = withoutReference(state, "2");

    expect(footnoteReferences(once.doc).map(({ id }) => id)).toEqual(["2"]);
    expect(footnoteStory(once, "2")).not.toBeNull();
    expect(footnoteStory(withoutReference(once, "2"), "2")).toBeNull();
  });

  it("brings back the reference and its footnote in one undo", () => {
    const state = opened();
    const deleted = withoutReference(state, "2");
    const undone = runCommand(deleted, undo);

    expect(footnoteReferences(undone.doc).map(({ id }) => id)).toEqual([
      "1",
      "2",
    ]);
    expect(
      sameStory(footnoteStory(undone, "2"), footnoteStory(state, "2"))
    ).toBe(true);
    expect(undoDepth(undone)).toBe(0);
    expect(footnoteStory(runCommand(undone, redo), "2")).toBeNull();
  });

  it("gives a copied reference a new id and a copy of its footnote without paragraph ids", () => {
    const state = opened();
    const copied = copiedBefore(state, "1", "One");
    const [copy, original] = footnoteReferences(copied.doc);
    const story = footnoteStory(copied, copy?.id ?? "");

    // Above the separator, both footnotes the text calls and the entry nobody calls
    expect(copy?.id).toBe("8");
    expect(copy?.node.attrs.referenceXml).toBeNull();
    expect(original?.id).toBe("1");
    expect(original?.node.attrs.referenceXml).toBe(
      '<w:footnoteReference w:id="1"/>'
    );
    expect(storyText(story)).toBe(storyText(footnoteStory(state, "1")));
    expect(story?.child(0).attrs.pAttrs).toBeNull();
    expect(story?.child(0).attrs.srcId).toBeNull();
    expect(footnoteStory(state, "1")?.child(0).attrs.pAttrs).toBe(
      'w14:paraId="10000001"'
    );
    expect(
      sameStory(footnoteStory(copied, "1"), footnoteStory(state, "1"))
    ).toBe(true);
    expect(undoDepth(copied)).toBe(1);
    expect(footnoteKeys(runCommand(copied, undo))).toEqual(footnoteKeys(state));
  });

  it("leaves a dragged reference its id and its footnote", () => {
    const state = opened();
    const { pos, node } = referenceTo(state, "1");
    const tr = state.tr.delete(pos, pos + 1);
    tr.insert(tr.mapping.map(rangeOfText(state.doc, "Three").from), node);
    const moved = state.apply(tr);

    expect(footnoteReferences(moved.doc).map(({ id }) => id)).toEqual([
      "2",
      "1",
    ]);
    expect(footnoteKeys(moved)).toEqual(footnoteKeys(state));
    expect(
      sameStory(footnoteStory(moved, "1"), footnoteStory(state, "1"))
    ).toBe(true);
  });

  it("keeps an entry the file arrived with no reference to", () => {
    const state = opened();
    const emptied = withoutReference(withoutReference(state, "1"), "2");

    expect(footnoteKeys(emptied)).toEqual(["footnote:-1", "footnote:7"]);
  });

  it("keeps a separator entry when the reference naming it is deleted", () => {
    const state = opened(
      `<w:p>${text("One")}${footnote("-1")}${footnote("1")}</w:p>`
    );
    const emptied = withoutReference(state, "-1");

    expect(footnoteReferences(emptied.doc).map(({ id }) => id)).toEqual(["1"]);
    expect(footnoteStory(emptied, "-1")).not.toBeNull();
  });

  it("leaves every footnote reference naming a story after each edit of the battery", () => {
    const edits: readonly {
      readonly name: string;
      edit(state: EditorState): EditorState;
    }[] = [
      {
        name: "delete a reference",
        edit: (state) => withoutReference(state, "2"),
      },
      {
        name: "copy a reference into another paragraph",
        edit: (state) => copiedBefore(state, "1", "Three"),
      },
      {
        name: "copy a paragraph holding a reference to the end",
        edit: (state) =>
          state.apply(
            state.tr.insert(state.doc.content.size, state.doc.child(0))
          ),
      },
      {
        name: "move a reference",
        edit: (state) => {
          const { pos, node } = referenceTo(state, "1");
          const tr = state.tr.delete(pos, pos + 1);
          tr.insert(tr.mapping.map(rangeOfText(state.doc, " two").to), node);
          return state.apply(tr);
        },
      },
      {
        name: "type over a selection running across a reference",
        edit: (state) => {
          const { pos } = referenceTo(state, "1");
          return state.apply(state.tr.insertText("typed", pos - 1, pos + 1));
        },
      },
      {
        name: "join two paragraphs",
        edit: (state) => {
          const end = state.doc.child(0).nodeSize;
          return state.apply(state.tr.delete(end - 1, end + 1));
        },
      },
      { name: "undo", edit: (state) => runCommand(state, undo) },
      { name: "redo", edit: (state) => runCommand(state, redo) },
      {
        name: "insert a footnote",
        edit: (state) => runCommand(select(state, 1), insertFootnote),
      },
      {
        name: "delete everything",
        edit: (state) =>
          state.apply(state.tr.delete(0, state.doc.content.size)),
      },
      { name: "undo that", edit: (state) => runCommand(state, undo) },
    ];

    edits.reduce((state, { name, edit }) => {
      const next = edit(state);
      const ids = footnoteReferences(next.doc).map(({ id }) => id);

      expect(next.doc.eq(state.doc), `${name} changed nothing`).toBe(false);
      expect(unnamed(next), `${name} left a reference naming no story`).toEqual(
        []
      );
      expect(new Set(ids).size, `${name} left two references to one id`).toBe(
        ids.length
      );
      return next;
    }, opened());
  });

  it("writes the file without the entry of a footnote whose reference was deleted", () => {
    const opening = importDocx(footnotesDocx(BODY, NOTES));
    const deleted = withoutReference(editorStateForSession(opening), "2");
    const exported = exportDocx(deleted.doc, opening.session);
    const part = decode(unzipSync(exported)["word/footnotes.xml"]);

    expect(part).not.toContain('w:id="2"');
    expect(part).toContain('<w:footnote w:id="1">');
    expect(part).toContain('<w:footnote w:id="7">');
    expect(part).toContain('w:type="separator"');
    expect(
      documentNotes(editorStateForSession(importDocx(exported))).map(
        ({ kind, id }) => `${kind}:${id}`
      )
    ).toEqual(["footnote:1"]);
  });

  it("writes a copied reference's footnote as an entry of its own", () => {
    const opening = importDocx(footnotesDocx(BODY, NOTES));
    const copied = copiedBefore(editorStateForSession(opening), "1", "One");
    const exported = exportDocx(copied.doc, opening.session);
    const part = decode(unzipSync(exported)["word/footnotes.xml"]);
    const reopened = editorStateForSession(importDocx(exported));

    expect(part.match(/10000001/g)).toHaveLength(1);
    expect(
      documentNotes(reopened).map(({ id, label, text }) => [id, label, text])
    ).toEqual([
      ["8", "1", " First note"],
      ["1", "2", " First note"],
      ["2", "3", " Second note"],
    ]);
  });
});
