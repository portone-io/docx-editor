// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it, vi } from "vitest";
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
import { type ProtectionState, protectionAllows } from "./protection";

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
 * The whole-document judgement is handed in watched, so the test can say it was never reached for,
 * which is the point of the fast path: a document of any size is typed in at the cost of the
 * stretch typed.
 */
describe("the guard's fast path", () => {
  const rules = (
    protection: ProtectionState["protection"]
  ): ProtectionState => ({
    protection,
    authorId: "me",
    editableComments: "own",
  });

  it("settles typing under none without the whole-document judgement", () => {
    const judge = vi.fn(protectionAllows);
    expect(
      protectionAllowsTransaction(typing(opened()), rules("none"), judge)
    ).toBe(true);
    expect(judge).not.toHaveBeenCalled();
  });

  it("refuses typing under comments without it either", () => {
    const judge = vi.fn(protectionAllows);
    expect(
      protectionAllowsTransaction(typing(opened()), rules("comments"), judge)
    ).toBe(false);
    expect(judge).not.toHaveBeenCalled();
  });

  it("reaches for it when a comment is touched", () => {
    const state = opened("comments");
    let tr: ReturnType<EditorState["tr"]["setMeta"]> | null = null;
    setCommentResolved("0", true)(state, (built) => (tr = built));
    if (tr === null) throw new Error("no transaction");
    const judge = vi.fn(protectionAllows);
    expect(protectionAllowsTransaction(tr, rules("comments"), judge)).toBe(
      true
    );
    expect(judge).toHaveBeenCalledTimes(1);
  });
});
