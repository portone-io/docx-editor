// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx, makeNotesDocx } from "../../__testing__/docx";
import { rangeOfText, runCommand } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { docxSchema } from "../../schema";
import {
  addComment,
  canEditComment,
  documentComments,
  selectComment,
} from "../commands/commentCommands";
import { documentNotes, noteProjection } from "../commands/noteQueries";
import { createEditorState } from "../createEditor";
import { commentProjection } from "./commentDecorations";
import { documentProjection } from "./documentProjection";

/** A document holding one paragraph of "alphabet", so an edit inside it is one character wide */
function paragraph(): PMNode {
  return docxSchema.node("doc", null, [
    docxSchema.node("paragraph", null, [docxSchema.text("alphabet")]),
  ]);
}

/** A projection over the document's text that reports how many times it was worked out */
function counted(name: string): {
  projection: ReturnType<typeof documentProjection<string>>;
  derivations: () => number;
} {
  let derivations = 0;
  const projection = documentProjection<string>(name, (doc) => {
    derivations += 1;
    return doc.textContent;
  });
  return { projection, derivations: () => derivations };
}

function caretAt(state: EditorState, pos: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, pos))
  );
}

describe("a value projected from the document", () => {
  it("derives once per document change and not per selection change", () => {
    const { projection, derivations } = counted("documentProjectionTestCount");
    let state = EditorState.create({
      doc: paragraph(),
      plugins: [projection.plugin],
    });
    expect(derivations()).toBe(1);

    // A caret moving changes nothing the value is worked out from
    state = caretAt(state, 3);
    expect(derivations()).toBe(1);

    // And reading it is not working it out again, however often the reader asks
    expect(projection.read(state)).toBe("alphabet");
    expect(projection.read(state)).toBe("alphabet");
    expect(derivations()).toBe(1);

    state = state.apply(state.tr.insertText("x", 1, 1));
    expect(projection.read(state)).toBe("xalphabet");
    expect(derivations()).toBe(2);
  });

  it("read answers for a state built without the plugin", () => {
    const { projection, derivations } = counted("documentProjectionTestBare");
    const state = EditorState.create({ doc: paragraph() });

    expect(projection.read(state)).toBe("alphabet");
    // Nothing holds it, so each reader of such a state works it out for itself
    expect(derivations()).toBe(1);
    expect(projection.read(state)).toBe("alphabet");
    expect(derivations()).toBe(2);
  });

  it("holds the comments of an editor state, and the ranges drawn over them", () => {
    const opened = createEditorState(
      importDocx(
        makeDocx(
          '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
        )
      ).doc
    );
    const at = rangeOfText(opened.doc, "beta");
    const commented = runCommand(
      opened,
      addComment(
        { text: "Look again", author: "Ada", date: "2026-01-01T00:00:00.000Z" },
        at
      )
    );

    const comments = commentProjection.read(commented).comments;
    expect(comments.map((comment) => comment.text)).toEqual(["Look again"]);
    const drawn = commentProjection
      .read(commented)
      .decorations.find()
      .map((decoration) => [decoration.from, decoration.to]);
    expect(drawn).toEqual([[comments[0].from, comments[0].to]]);
    const plugin = commentProjection.plugin;
    expect(plugin.props.decorations?.call(plugin, commented)).toBe(
      commentProjection.read(commented).decorations
    );

    // The list and the ranges the editor holds are the ones its last edit worked out, so a caret
    // moving through the document leaves both standing rather than walking it again
    const moved = caretAt(commented, 1);
    expect(commentProjection.read(moved).comments).toBe(comments);
    expect(commentProjection.read(moved).decorations).toBe(
      commentProjection.read(commented).decorations
    );
  });

  it("holds the notes of an editor state", () => {
    const opened = createEditorState(importDocx(makeNotesDocx()).doc);

    const notes = noteProjection.read(opened);
    expect(notes.map((note) => note.kind)).toEqual(["footnote", "endnote"]);
    expect(noteProjection.read(caretAt(opened, 1))).toBe(notes);
  });
});

describe("public projection results", () => {
  function commentedState() {
    return createEditorState(
      docxSchema.node("doc", null, [
        docxSchema.node("paragraph", null, [
          docxSchema.node("commentStart", { id: "0" }),
          docxSchema.text("Alpha"),
          docxSchema.node("commentEnd", { id: "0" }),
          docxSchema.node("commentReference", {
            id: "0",
            text: "Original",
            authorId: "other",
            replies: [
              {
                id: "1",
                author: "Other",
                authorId: "other",
                initials: null,
                date: null,
                text: "Reply",
              },
            ],
          }),
        ]),
      ]),
      { author: { id: "me", name: "Me" }, editableComments: "own" }
    );
  }

  it("keeps caller edits to comment and reply records out of ownership checks", () => {
    const state = commentedState();
    const comments = documentComments(state);
    comments[0].authorId = "me";
    comments[0].text = "Locally formatted";
    comments[0].replies[0].authorId = "me";
    comments[0].replies[0].text = "Locally formatted reply";
    expect(canEditComment(state, "0")).toBe(false);
    expect(canEditComment(state, "0", "1")).toBe(false);
    expect(documentComments(state)[0]).toMatchObject({
      text: "Original",
      replies: [{ text: "Reply" }],
    });
  });

  it("keeps caller edits to a comment's range out of anchor selection", () => {
    const state = commentedState();
    const comment = documentComments(state)[0];
    const originalRange = { from: comment.from, to: comment.to };
    comment.from = state.doc.content.size + 100;
    const selected = runCommand(state, selectComment("0"));
    expect(selected.selection).toMatchObject(originalRange);
  });

  it("keeps caller edits to note records out of later query results", () => {
    const state = createEditorState(importDocx(makeNotesDocx()).doc);
    const expected = { ...documentNotes(state)[0] };
    documentNotes(state)[0].text = "Locally formatted";
    expect(documentNotes(state)[0]).toEqual(expected);
  });

  it("uses the first reference with an id and still refuses a missing id", () => {
    const state = createEditorState(
      docxSchema.node("doc", null, [
        docxSchema.node("paragraph", null, [
          docxSchema.node("commentReference", { id: "0", authorId: "me" }),
          docxSchema.text("Alpha"),
          docxSchema.node("commentReference", { id: "0", authorId: "other" }),
        ]),
      ]),
      { author: { id: "me", name: "Me" }, editableComments: "own" }
    );
    expect(canEditComment(state, "0")).toBe(true);
    expect(runCommand(state, selectComment("0")).selection.from).toBe(1);
    expect(canEditComment(state, "missing")).toBe(false);
    expect(selectComment("missing")(state)).toBe(false);
  });
});
