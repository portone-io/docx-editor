// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
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
import { displayOnly } from "./displayDerivation";
import {
  protectionAllowsTransaction,
  transactionAllowed,
  transactionTouchesComments,
} from "./guards";
import type { EditingProtection, ProtectionState } from "./protection";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const COMMENTED =
  `<w:p>${run("Alpha ")}` +
  '<w:commentRangeStart w:id="0"/>' +
  run("beta") +
  '<w:commentRangeEnd w:id="0"/>' +
  '<w:r><w:commentReference w:id="0"/></w:r>' +
  `${run(" gamma")}</w:p><w:p>${run("Delta")}</w:p>`;

const lockedPr =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const cell = (text: string) => `<w:tc><w:p>${run(text)}</w:p></w:tc>`;

/** A two-column table whose left cell stands inside a control that shuts it */
const LOCKED_CELL =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr><w:sdt>${lockedPr}<w:sdtContent>${cell("Locked")}</w:sdtContent></w:sdt>` +
  `${cell("Open")}</w:tr></w:tbl>`;

function opened(
  protection: EditingProtection = "none",
  body = COMMENTED
): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc, {
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

/** The first block of this type reading exactly this text, and where it stands */
function blockHolding(
  doc: PMNode,
  typeName: string,
  text: string
): { pos: number; node: PMNode } {
  const found: { pos: number; node: PMNode }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === typeName && node.textContent === text) {
      found.push({ pos, node });
    }
    return true;
  });
  const first = found[0];
  if (first === undefined) throw new Error(`no ${typeName} holding ${text}`);
  return first;
}

/**
 * What a node draws with is worked out from its source (`./attrRoles`), so a transaction that
 * only writes those values again is no edit, and it carries the `displayOnly` pass to say so. The
 * pass is a claim the steps have to bear out: a step touching anything the role table calls
 * source, a lock flag included, is refused the claim and judged as any edit.
 */
describe("the display-only pass", () => {
  const CENTERED = { format: { align: "center" } };

  function withLockedCell(protection: EditingProtection = "none"): EditorState {
    return opened(protection, COMMENTED + LOCKED_CELL);
  }

  /** The paragraph inside the locked cell, rewritten around its content with these attrs */
  function rewritten(
    state: EditorState,
    attrs: Record<string, unknown>
  ): Transaction {
    const { pos, node } = blockHolding(state.doc, "paragraph", "Locked");
    return state.tr.setNodeMarkup(pos, null, { ...node.attrs, ...attrs });
  }

  it("lets a display value written again inside a locked cell through", () => {
    const state = withLockedCell();
    const tr = rewritten(state, CENTERED).setMeta(displayOnly, true);

    expect(transactionAllowed(tr, state)).toBe(true);
    expect(state.apply(tr).doc.eq(state.doc)).toBe(false);
  });

  it("judges the same change as an edit when nothing claims it is a re-derivation", () => {
    const state = withLockedCell();
    expect(transactionAllowed(rewritten(state, CENTERED), state)).toBe(false);
  });

  it("refuses the claim to a step that rewrites a source attr", () => {
    const state = withLockedCell();
    const tr = rewritten(state, {
      pPr: '<w:pPr><w:jc w:val="center"/></w:pPr>',
      ...CENTERED,
    }).setMeta(displayOnly, true);

    expect(transactionAllowed(tr, state)).toBe(false);
  });

  it("refuses the claim to a step that lifts a cell's lock, however it is written", () => {
    const state = withLockedCell();
    const { pos, node } = blockHolding(state.doc, "tableCell", "Locked");
    const whole = state.tr
      .setNodeMarkup(pos, null, { ...node.attrs, sdtContentsLocked: false })
      .setMeta(displayOnly, true);
    const oneAttr = state.tr
      .setNodeAttribute(pos, "sdtContentsLocked", false)
      .setMeta(displayOnly, true);

    expect(node.attrs.sdtContentsLocked).toBe(true);
    expect(transactionAllowed(whole, state)).toBe(false);
    expect(transactionAllowed(oneAttr, state)).toBe(false);
  });

  it("refuses the claim to a transaction with one real edit among its steps", () => {
    const state = withLockedCell();
    const tr = rewritten(state, CENTERED)
      .insertText("x", rangeOfText(state.doc, "Locked").from)
      .setMeta(displayOnly, true);

    expect(transactionAllowed(tr, state)).toBe(false);
  });

  it("is no content change under comments and no edit under readOnly", () => {
    for (const protection of ["comments", "readOnly"] as const) {
      const state = withLockedCell(protection);
      expect(
        transactionAllowed(
          rewritten(state, CENTERED).setMeta(displayOnly, true),
          state
        )
      ).toBe(true);
      expect(
        transactionAllowed(typing(state).setMeta(displayOnly, true), state)
      ).toBe(false);
    }
  });
});
