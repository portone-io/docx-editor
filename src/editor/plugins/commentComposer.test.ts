// @vitest-environment jsdom
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { rangeOfText } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { createEditorState } from "../createEditor";
import {
  closeCommentComposer,
  commentComposerRange,
  isCommentComposerOpen,
  openCommentComposer,
} from "./commentComposer";
import { setProtection } from "./documentProtection";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const BODY =
  `<w:p>${run("Alpha beta")}</w:p>` + `<w:p>${run("Gamma delta")}</w:p>`;

function opened(): EditorState {
  return createEditorState(importDocx(makeDocx(BODY)).doc, {
    author: { id: "grace", name: "Grace" },
  });
}

function selecting(state: EditorState, text: string): EditorState {
  const { from, to } = rangeOfText(state.doc, text);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

/** Runs the command over the state and answers both what it said and the state it left */
function attempt(
  state: EditorState,
  command: Command
): { answered: boolean; state: EditorState } {
  let after = state;
  const answered = command(state, (tr) => {
    after = after.apply(tr);
  });
  return { answered, state: after };
}

describe("the comment composer's state", () => {
  it("opens only where a comment can go", () => {
    const caret = opened();
    // A caret marks no stretch of text, so there is nothing to comment on
    expect(attempt(caret, openCommentComposer).answered).toBe(false);
    expect(isCommentComposerOpen(caret)).toBe(false);

    const selected = selecting(caret, "beta");
    const { answered, state } = attempt(selected, openCommentComposer);
    expect(answered).toBe(true);
    expect(commentComposerRange(state)).toEqual(rangeOfText(state.doc, "beta"));

    // Closing answers false the second time, so a key press falls through to what is behind it
    const closed = attempt(state, closeCommentComposer);
    expect(closed.answered).toBe(true);
    expect(isCommentComposerOpen(closed.state)).toBe(false);
    expect(attempt(closed.state, closeCommentComposer).answered).toBe(false);
  });

  it("keeps the anchored range across an edit elsewhere", () => {
    const state = attempt(
      selecting(opened(), "beta"),
      openCommentComposer
    ).state;
    const anchored = rangeOfText(state.doc, "beta");

    // Text typed into the paragraph ahead of the anchor pushes the whole stretch along
    const edited = state.apply(state.tr.insertText("Long ", 1, 1));
    expect(commentComposerRange(edited)).toEqual({
      from: anchored.from + 5,
      to: anchored.to + 5,
    });
    expect(edited.doc.textBetween(anchored.from + 5, anchored.to + 5)).toBe(
      "beta"
    );

    // And an edit in the second paragraph, past the anchor, leaves it where it stands
    const later = rangeOfText(edited.doc, "delta");
    const further = edited.apply(edited.tr.insertText("x", later.from));
    expect(commentComposerRange(further)).toEqual({
      from: anchored.from + 5,
      to: anchored.to + 5,
    });
  });

  it("closes when the anchored text is deleted", () => {
    const state = attempt(
      selecting(opened(), "beta"),
      openCommentComposer
    ).state;
    const anchored = rangeOfText(state.doc, "beta");

    const deleted = state.apply(state.tr.delete(anchored.from, anchored.to));
    expect(commentComposerRange(deleted)).toBeNull();
  });

  it("closes when the protection turns readOnly", () => {
    const state = attempt(
      selecting(opened(), "beta"),
      openCommentComposer
    ).state;

    // The mode is switched on the open document, which changes no text at all
    const switching = setProtection(state.tr, {
      protection: "readOnly",
      author: null,
      editableComments: "own",
    });
    expect(switching.docChanged).toBe(false);
    expect(isCommentComposerOpen(state.apply(switching))).toBe(false);

    // A commenter is offered it, since writing comments is what that mode is for
    const commenting = state.apply(
      setProtection(state.tr, {
        protection: "comments",
        author: { id: "grace", name: "Grace" },
        editableComments: "own",
      })
    );
    expect(isCommentComposerOpen(commenting)).toBe(true);
  });
});
