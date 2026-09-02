// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";
import { DocAttrStep } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import {
  addComment,
  removeComment,
  setCommentResolved,
} from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import {
  protectionAllowsTransaction,
  transactionTouchesComments,
} from "./locks";
import type { ProtectionState } from "./protection";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const COMMENTED =
  `<w:p>${run("Alpha ")}` +
  '<w:commentRangeStart w:id="0"/>' +
  run("beta") +
  '<w:commentRangeEnd w:id="0"/>' +
  '<w:r><w:commentReference w:id="0"/></w:r>' +
  `${run(" gamma")}</w:p><w:p>${run("Delta")}</w:p>`;

function opened(
  protection: "none" | "comments" | "readOnly" = "none"
): EditorState {
  return createEditorState(importDocx(makeDocx(COMMENTED)).doc, {
    protection,
    author: { id: "me", name: "Me" },
  });
}

function typing(state: EditorState) {
  return state.tr.insertText("x", rangeOfText(state.doc, "Delta").from);
}

describe("transactionTouchesComments", () => {
  it("is false for typing, formatting and a selection away from every comment", () => {
    const state = opened();
    const delta = rangeOfText(state.doc, "Delta");
    expect(transactionTouchesComments(typing(state))).toBe(false);
    expect(
      transactionTouchesComments(
        state.tr.addMark(delta.from, delta.to, state.schema.marks.run.create())
      )
    ).toBe(false);
    expect(
      transactionTouchesComments(
        state.tr.setSelection(TextSelection.create(state.doc, delta.from))
      )
    ).toBe(false);
  });

  it("is true for a comment put in, taken out, settled, or deleted with its text", () => {
    const state = opened();
    const delta = rangeOfText(state.doc, "Delta");
    const selected = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, delta.from, delta.to)
      )
    );
    const built: ReturnType<EditorState["tr"]["setMeta"]>[] = [];
    addComment({ text: "n", author: "Me" })(selected, (tr) => built.push(tr));
    removeComment("0")(state, (tr) => built.push(tr));
    setCommentResolved("0", true)(state, (tr) => built.push(tr));
    expect(built).toHaveLength(3);
    for (const tr of built) expect(transactionTouchesComments(tr)).toBe(true);

    const beta = rangeOfText(state.doc, "beta");
    expect(
      transactionTouchesComments(state.tr.delete(beta.from - 1, beta.to + 2))
    ).toBe(true);
  });
});

/**
 * Whether the transaction touches a comment is what says whether the whole-document judgement is
 * reached for at all, which is the point of the fast path: a document of any size is typed in at
 * the cost of the stretch typed.
 */
describe("the guard's fast path", () => {
  const rules = (
    protection: ProtectionState["protection"]
  ): ProtectionState => ({
    protection,
    authorId: "me",
    editableComments: "own",
  });

  it("settles typing without the whole-document judgement, whichever way it goes", () => {
    const tr = typing(opened());
    expect(transactionTouchesComments(tr)).toBe(false);
    expect(protectionAllowsTransaction(tr, rules("none"))).toBe(true);
    expect(protectionAllowsTransaction(tr, rules("comments"))).toBe(false);
  });

  it("reaches for it when a comment is touched", () => {
    const state = opened("comments");
    let tr: ReturnType<EditorState["tr"]["setMeta"]> | null = null;
    setCommentResolved("0", true)(state, (built) => (tr = built));
    if (tr === null) throw new Error("no transaction");
    expect(transactionTouchesComments(tr)).toBe(true);
    expect(protectionAllowsTransaction(tr, rules("comments"))).toBe(true);
  });

  it("reaches for it over a step whose kind it does not know", () => {
    const state = opened("comments");
    const tr = state.tr.step(new DocAttrStep("unknown", 1));
    expect(transactionTouchesComments(tr)).toBe(true);
    // The judgement finds the document as it was and lets the step through. Answering that a step
    // of an unknown kind touches no comment would have refused it under comments instead
    expect(protectionAllowsTransaction(tr, rules("comments"))).toBe(true);
  });
});
