// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { undoDepth } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { decode, makeNotesDocx, NOTE_BODY } from "../../__testing__/docx";
import { rangeOfText, runCommand, select } from "../../__testing__/editing";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import { storyFromText, storyText } from "../../docx/story";
import { docxSchema } from "../../schema";
import type { EditingProtection } from "../../schema/protection";
import {
  type NoteKind,
  sameStory,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { editorStateForSession } from "../createEditor";
import { undo } from "./historyCommands";
import {
  canInsertFootnote,
  insertEndnote,
  insertFootnote,
  setEndnoteBody,
  setFootnoteBody,
} from "./noteCommands";
import { documentNotes } from "./noteQueries";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const footnote = (id: string) =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
const locked = (inner: string) =>
  '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
  `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const FOOTNOTE_STYLES =
  '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>' +
  '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';

interface Opening {
  readonly body?: string;
  readonly styles?: string | null;
  readonly footnotes?: string;
  readonly endnotes?: string;
  readonly protection?: EditingProtection;
}

function bytesOf({
  body = NOTE_BODY,
  styles = null,
  footnotes,
  endnotes,
}: Opening) {
  const bytes = makeNotesDocx(body, styles);
  if (footnotes === undefined && endnotes === undefined) return bytes;
  const parts = unzipSync(bytes);
  const encoder = new TextEncoder();
  if (footnotes !== undefined) {
    parts["word/footnotes.xml"] = encoder.encode(
      `<w:footnotes xmlns:w="${W_NS}">${footnotes}</w:footnotes>`
    );
  }
  if (endnotes !== undefined) {
    parts["word/endnotes.xml"] = encoder.encode(
      `<w:endnotes xmlns:w="${W_NS}">${endnotes}</w:endnotes>`
    );
  }
  return zipSync(parts);
}

function opened(opening: Opening = {}): EditorState {
  return editorStateForSession(importDocx(bytesOf(opening)), {
    protection: opening.protection ?? "none",
  });
}

function noteIds(doc: PMNode, kind: NoteKind): string[] {
  const ids: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "noteReference" && node.attrs.kind === kind) {
      ids.push(String(node.attrs.id));
    }
    return true;
  });
  return ids;
}

function footnoteIds(doc: PMNode): string[] {
  return noteIds(doc, "footnote");
}

function referenceNode(doc: PMNode, id: string): PMNode {
  let found: PMNode | null = null;
  doc.descendants((node) => {
    if (
      found === null &&
      node.type.name === "noteReference" &&
      node.attrs.id === id
    ) {
      found = node;
    }
    return found === null;
  });
  if (found === null) throw new Error(`no reference to ${id}`);
  return found;
}

function runProps(node: PMNode | null | undefined): unknown {
  return node?.marks.find((mark) => mark.type.name === "run")?.attrs.rPr;
}

function footnoteStory(state: EditorState, id: string): PMNode | null {
  return storyNodeOf(state.doc, storyKey("footnote", id));
}

/** Runs the command for real, and reports what it answered and whether the state moved */
function attempt(
  state: EditorState,
  command: Command
): { answered: boolean; changed: boolean } {
  let after = state;
  const answered = command(state, (tr) => {
    after = after.apply(tr);
  });
  return { answered, changed: after !== state };
}

const REFUSED = { answered: false, changed: false };

/** A footnote body a composer of one's own would write: a bold run and a second paragraph */
function formattedBody(): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create(
      null,
      docxSchema.text("Bold words", [
        docxSchema.marks.run.create({ rPr: "<w:rPr><w:b/></w:rPr>" }),
      ])
    ),
    docxSchema.nodes.paragraph.create(null, docxSchema.text("Second line")),
  ]);
}

describe("insertFootnote", () => {
  it("inserts a footnote reference at the end of the selection with a new footnote", () => {
    const state = opened();
    const { from, to } = rangeOfText(state.doc, "and more");
    const after = runCommand(select(state, from, to), insertFootnote);
    const story = footnoteStory(after, "3");

    expect(footnoteIds(after.doc)).toEqual(["2", "3"]);
    expect(after.doc.nodeAt(to)?.attrs.id).toBe("3");
    expect(after.doc.textBetween(from, to)).toBe("and more");
    expect(after.selection.from).toBe(from);
    expect(story?.childCount).toBe(1);
    expect(story?.child(0).firstChild?.attrs.element).toBe("footnoteRef");
    expect(storyText(story)).toBe("");
    expect(
      documentNotes(after).map(({ kind, id, label }) => [kind, id, label])
    ).toEqual([
      ["footnote", "2", "1"],
      ["footnote", "3", "2"],
      ["endnote", "3", "1"],
    ]);
    expect(undoDepth(after)).toBe(1);
    expect(footnoteStory(runCommand(after, undo), "3")).toBeNull();
  });

  it("takes an id above every entry the part holds, separators included", () => {
    const state = opened({
      body: `<w:p>${text("Text")}${footnote("2")}${footnote("9")}</w:p>`,
      footnotes:
        '<w:footnote w:type="separator" w:id="12"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        `<w:footnote w:id="2"><w:p>${text("Called")}</w:p></w:footnote>` +
        `<w:footnote w:id="5"><w:p>${text("Never called")}</w:p></w:footnote>`,
    });
    const after = runCommand(select(state, 1), insertFootnote);

    expect(footnoteIds(after.doc)).toEqual(["13", "2", "9"]);
  });

  it("takes no id of a footnote deleted since the document was opened", () => {
    const state = opened();
    const at = rangeOfText(state.doc, "Text").to;
    const emptied = state.apply(state.tr.delete(at, at + 1));
    const after = runCommand(select(emptied, 1), insertFootnote);

    expect(footnoteStory(emptied, "2")).toBeNull();
    expect(footnoteIds(after.doc)).toEqual(["3"]);
  });

  it("writes the footnote text and reference styles the document defines", () => {
    const opening = importDocx(bytesOf({ styles: FOOTNOTE_STYLES }));
    const state = editorStateForSession(opening);
    const at = rangeOfText(state.doc, "Text").to;
    const after = runCommand(select(state, at), insertFootnote);
    const paragraph = footnoteStory(after, "3")?.child(0);
    const reference = '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>';

    expect(runProps(referenceNode(after.doc, "3"))).toBe(reference);
    expect(paragraph?.attrs.pPr).toBe(
      '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
    );
    expect(runProps(paragraph?.firstChild)).toBe(reference);

    const parts = unzipSync(exportDocx(after.doc, opening.session));
    expect(decode(parts["word/document.xml"])).toContain(
      `<w:r>${reference}<w:footnoteReference w:id="3"/></w:r>`
    );
    expect(decode(parts["word/footnotes.xml"])).toContain(
      `<w:footnote w:id="3"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r>${reference}<w:footnoteRef/></w:r></w:p></w:footnote>`
    );
  });

  it("marks the reference superscript where the document defines no footnote styles", () => {
    const state = opened();
    const after = runCommand(select(state, 1), insertFootnote);
    const superscript = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>';
    const paragraph = footnoteStory(after, "3")?.child(0);

    expect(runProps(referenceNode(after.doc, "3"))).toBe(superscript);
    expect(runProps(paragraph?.firstChild)).toBe(superscript);
    expect(paragraph?.attrs.pPr).toBeNull();
  });

  it("cannot insert a footnote where the caret stands in locked content", () => {
    const state = opened({
      body: `<w:p>${text("open")}${locked(text("shut"))}</w:p>`,
    });
    const inside = select(state, rangeOfText(state.doc, "shut").from + 1);

    expect(canInsertFootnote(inside)).toBe(false);
    expect(attempt(inside, insertFootnote)).toEqual(REFUSED);
    expect(
      canInsertFootnote(select(state, rangeOfText(state.doc, "open").to))
    ).toBe(true);
  });

  it.each(["readOnly", "comments"] as const)(
    "cannot insert a footnote under the %s protection",
    (protection) => {
      const state = select(opened({ protection }), 1);

      expect(canInsertFootnote(state)).toBe(false);
      expect(attempt(state, insertFootnote)).toEqual(REFUSED);
    }
  );
});

describe("setFootnoteBody", () => {
  it("writes a formatted body into a footnote in one undoable step", () => {
    const state = opened();
    const after = runCommand(state, setFootnoteBody("2", formattedBody()));

    expect(sameStory(footnoteStory(after, "2"), formattedBody())).toBe(true);
    expect(documentNotes(after)[0]?.text).toBe("Bold words\nSecond line");
    expect(undoDepth(after)).toBe(1);
    expect(
      sameStory(
        footnoteStory(runCommand(after, undo), "2"),
        footnoteStory(state, "2")
      )
    ).toBe(true);
  });

  it("refuses a body for a footnote the document does not refer to", () => {
    const state = opened({
      footnotes:
        `<w:footnote w:id="2"><w:p>${text("Called")}</w:p></w:footnote>` +
        `<w:footnote w:id="5"><w:p>${text("Never called")}</w:p></w:footnote>`,
    });

    expect(footnoteStory(state, "5")).not.toBeNull();
    expect(attempt(state, setFootnoteBody("5", formattedBody()))).toEqual(
      REFUSED
    );
    expect(attempt(state, setFootnoteBody("99", formattedBody()))).toEqual(
      REFUSED
    );
    // An endnote is no footnote, whatever id it carries
    expect(attempt(state, setFootnoteBody("3", formattedBody()))).toEqual(
      REFUSED
    );
  });

  it("refuses a body for a separator entry", () => {
    const state = opened({
      body: `<w:p>${text("Text")}${footnote("-1")}${footnote("2")}</w:p>`,
    });

    expect(attempt(state, setFootnoteBody("-1", formattedBody()))).toEqual(
      REFUSED
    );
  });

  it("leaves the history alone for a body that says the same thing", () => {
    const state = opened();
    const held = footnoteStory(state, "2");
    if (held === null) throw new Error("no footnote 2");
    const rebuilt = importDocx(bytesOf({})).doc;
    const alike = storyNodeOf(rebuilt, storyKey("footnote", "2"));
    if (alike === null) throw new Error("no footnote 2 in the reopened file");

    expect(attempt(state, setFootnoteBody("2", held))).toEqual(REFUSED);
    // Read from another opening, so every block names another source
    expect(attempt(state, setFootnoteBody("2", alike))).toEqual(REFUSED);
    expect(undoDepth(state)).toBe(0);
  });

  it("refuses a footnote body edit under comment mode", () => {
    const state = opened({ protection: "comments" });

    expect(
      attempt(state, setFootnoteBody("2", storyFromText("Rewritten")))
    ).toEqual(REFUSED);
  });

  it("refuses a footnote body edit whose reference stands in locked content", () => {
    const state = opened({
      body:
        `<w:p>${text("open")}${footnote("2")}</w:p>` +
        `<w:p>${locked(`${text("shut")}${footnote("4")}`)}</w:p>`,
      footnotes:
        `<w:footnote w:id="2"><w:p>${text("Open note")}</w:p></w:footnote>` +
        `<w:footnote w:id="4"><w:p>${text("Shut note")}</w:p></w:footnote>`,
    });

    expect(attempt(state, setFootnoteBody("4", formattedBody()))).toEqual(
      REFUSED
    );
    expect(attempt(state, setFootnoteBody("2", formattedBody()))).toEqual({
      answered: true,
      changed: true,
    });
  });
});

describe("insertEndnote", () => {
  it("inserts an endnote reference at the end of the selection with a new endnote", () => {
    const state = opened();
    const { from, to } = rangeOfText(state.doc, "and more");
    const after = runCommand(select(state, from, to), insertEndnote);
    const story = storyNodeOf(after.doc, storyKey("endnote", "4"));

    expect(noteIds(after.doc, "endnote")).toEqual(["4", "3"]);
    expect(after.doc.nodeAt(to)?.attrs.id).toBe("4");
    expect(after.doc.textBetween(from, to)).toBe("and more");
    expect(after.selection.from).toBe(from);
    expect(story?.child(0).firstChild?.attrs.element).toBe("endnoteRef");
    expect(storyText(story)).toBe("");
    expect(
      documentNotes(after).map(({ kind, id, label }) => [kind, id, label])
    ).toEqual([
      ["footnote", "2", "1"],
      ["endnote", "4", "1"],
      ["endnote", "3", "2"],
    ]);
    expect(undoDepth(after)).toBe(1);
    expect(
      storyNodeOf(runCommand(after, undo).doc, storyKey("endnote", "4"))
    ).toBeNull();
  });

  it("takes an endnote id above every entry the endnotes part holds, separators included", () => {
    const state = opened({
      body: `<w:p>${text("Text")}<w:r><w:endnoteReference w:id="2"/></w:r></w:p>`,
      endnotes:
        '<w:endnote w:type="separator" w:id="12"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
        `<w:endnote w:id="2"><w:p>${text("Called")}</w:p></w:endnote>`,
    });
    const after = runCommand(select(state, 1), insertEndnote);

    expect(noteIds(after.doc, "endnote")).toEqual(["13", "2"]);
    // The footnote ids are the endnotes part's own business and are counted apart from it
    expect(footnoteIds(after.doc)).toEqual([]);
  });

  it("refuses an endnote body edit under comment mode", () => {
    const state = opened({ protection: "comments" });

    expect(
      attempt(state, setEndnoteBody("3", storyFromText("Rewritten")))
    ).toEqual(REFUSED);
  });

  it("writes a formatted body into an endnote in one undoable step", () => {
    const state = opened();
    const body = formattedBody();
    const after = runCommand(state, setEndnoteBody("3", body));

    expect(
      sameStory(storyNodeOf(after.doc, storyKey("endnote", "3")), body)
    ).toBe(true);
    expect(undoDepth(after)).toBe(1);
  });
});
