// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import {
  addComment,
  addCommentReply,
  removeComment,
  removeCommentReply,
  setCommentResolved,
  updateComment,
  updateCommentReply,
} from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import {
  changesOnlyComments,
  commentAdditionsBy,
  isCommentNode,
  type ProtectionState,
  protectionAllows,
} from "./protection";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const BODY =
  `<w:p>${run("Alpha ")}${run("beta")}</w:p>` + `<w:p>${run("Gamma")}</w:p>`;

/** A state with nothing shut, anyone's comment included, so that every document shape below can be built through the commands */
function stateOf(doc: PMNode): EditorState {
  return createEditorState(doc, { editableComments: "all" });
}

function opened(): EditorState {
  return stateOf(importDocx(makeDocx(BODY)).doc);
}

function applied(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
}

function selecting(state: EditorState, text: string): EditorState {
  const { from, to } = rangeOfText(state.doc, text);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

/** The document with one comment on "beta" written under this identity, or under none */
function commented(authorId: string | null): EditorState {
  return applied(
    selecting(opened(), "beta"),
    addComment({
      text: "note",
      author: "Ada",
      authorId: authorId ?? undefined,
    })
  );
}

const commentId = (state: EditorState): string => {
  let id: string | null = null;
  state.doc.descendants((node) => {
    if (
      node.type.name === "commentReference" &&
      typeof node.attrs.id === "string"
    ) {
      id = node.attrs.id;
    }
  });
  if (id === null) throw new Error("no comment in the document");
  return id;
};

/** Where the one comment reference stands, which is the node every attribute of a comment travels on */
function referenceAt(doc: PMNode): { pos: number; node: PMNode } {
  let found: { pos: number; node: PMNode } | null = null;
  doc.descendants((node, pos) => {
    if (found === null && node.type.name === "commentReference") {
      found = { pos, node };
    }
    return true;
  });
  if (found === null) throw new Error("no comment in the document");
  return found;
}

/**
 * The document with the attributes of its one comment written over, which is the shape a step a
 * consumer brought could take, and the shape a file handed back can simply claim.
 */
function rewritten(state: EditorState, attrs: Record<string, unknown>): PMNode {
  const { pos, node } = referenceAt(state.doc);
  // The document the step leaves behind rather than the state it would leave, since the guard
  // itself refuses the step: this is a document arriving from outside the editor
  return state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs })
    .doc;
}

/** The document with the comment's three markers lifted off "beta" and put around "Gamma" */
function moved(state: EditorState): PMNode {
  const markers: { pos: number; node: PMNode }[] = [];
  state.doc.descendants((node, pos) => {
    if (isCommentNode(node)) markers.push({ pos, node });
    return true;
  });
  const tr = state.tr;
  for (const { pos, node } of [...markers].reverse()) {
    tr.delete(pos, pos + node.nodeSize);
  }
  const target = rangeOfText(tr.doc, "Gamma");
  const [start, end, reference] = markers.map((marker) => marker.node);
  tr.insert(target.to, [end, reference]);
  tr.insert(target.from, start);
  return tr.doc;
}

function typed(state: EditorState): PMNode {
  return state.apply(state.tr.insertText("x", 1)).doc;
}

const rules = (
  protection: ProtectionState["protection"],
  authorId: string | null = "me",
  editableComments: ProtectionState["editableComments"] = "own"
): ProtectionState => ({ protection, authorId, editableComments });

describe("changesOnlyComments", () => {
  it("holds for an unchanged document", () => {
    const state = opened();
    expect(changesOnlyComments(state.doc, state.doc)).toBe(true);
  });

  it("holds for a comment put on text and taken off it again", () => {
    const before = opened();
    const after = commented("me");
    expect(changesOnlyComments(before.doc, after.doc)).toBe(true);
    expect(changesOnlyComments(after.doc, before.doc)).toBe(true);
    const removed = applied(after, removeComment(commentId(after)));
    expect(changesOnlyComments(after.doc, removed.doc)).toBe(true);
  });

  it("holds for a reply, a resolution and a rewritten comment", () => {
    const state = commented(null);
    const id = commentId(state);
    for (const command of [
      addCommentReply(id, { text: "reply", author: "Bo" }),
      setCommentResolved(id, true),
      updateComment(id, "rewritten"),
    ]) {
      expect(changesOnlyComments(state.doc, applied(state, command).doc)).toBe(
        true
      );
    }
  });

  it("holds for a comment carried onto other text, which ownership answers for instead", () => {
    const state = commented("other");
    expect(changesOnlyComments(state.doc, moved(state))).toBe(true);
  });

  it("does not hold for typed text, with or without a comment beside it", () => {
    const plain = opened();
    expect(changesOnlyComments(plain.doc, typed(plain))).toBe(false);
    const state = commented(null);
    expect(changesOnlyComments(state.doc, typed(state))).toBe(false);
    expect(changesOnlyComments(plain.doc, typed(state))).toBe(false);
  });
});

describe("protectionAllows", () => {
  it("refuses every change under readOnly", () => {
    const state = opened();
    expect(protectionAllows(state.doc, typed(state), rules("readOnly"))).toBe(
      false
    );
    expect(
      protectionAllows(state.doc, commented("me").doc, rules("readOnly"))
    ).toBe(false);
  });

  it("takes a comment and refuses text under comments", () => {
    const state = opened();
    expect(
      protectionAllows(state.doc, commented("me").doc, rules("comments"))
    ).toBe(true);
    expect(protectionAllows(state.doc, typed(state), rules("comments"))).toBe(
      false
    );
  });

  it("takes text under none", () => {
    const state = opened();
    expect(protectionAllows(state.doc, typed(state), rules("none"))).toBe(true);
  });

  describe("over another author's comment", () => {
    const state = commented("other");
    const id = commentId(state);

    it.each([
      ["rewriting it", updateComment(id, "rewritten")],
      ["deleting it", removeComment(id)],
    ] as const)("refuses %s under own", (_name, command) => {
      const after = applied(state, command).doc;
      expect(protectionAllows(state.doc, after, rules("comments"))).toBe(false);
      expect(protectionAllows(state.doc, after, rules("none"))).toBe(false);
    });

    it.each([
      ["rewriting it", updateComment(id, "rewritten")],
      ["deleting it", removeComment(id)],
    ] as const)("takes %s under all", (_name, command) => {
      const after = applied(state, command).doc;
      expect(
        protectionAllows(state.doc, after, rules("comments", "me", "all"))
      ).toBe(true);
    });

    it.each([
      ["replying to it", addCommentReply(id, { text: "reply", author: "Me" })],
      ["resolving it", setCommentResolved(id, true)],
    ] as const)("takes %s from anyone", (_name, command) => {
      const after = applied(state, command).doc;
      expect(protectionAllows(state.doc, after, rules("comments"))).toBe(true);
    });

    it("refuses carrying it onto other text under own and takes it under all", () => {
      const after = moved(state);
      expect(protectionAllows(state.doc, after, rules("comments"))).toBe(false);
      expect(protectionAllows(state.doc, after, rules("none"))).toBe(false);
      expect(
        protectionAllows(state.doc, after, rules("comments", "me", "all"))
      ).toBe(true);
    });

    it("takes a body edit that sweeps it away under none", () => {
      const from = rangeOfText(state.doc, "beta").from - 1;
      const swept = state.apply(
        state.tr.delete(from, state.doc.child(0).nodeSize - 1)
      ).doc;
      expect(changesOnlyComments(state.doc, swept)).toBe(false);
      expect(protectionAllows(state.doc, swept, rules("none"))).toBe(true);
    });
  });

  describe("over one's own comment and a comment nobody claimed", () => {
    it.each([
      ["one's own", "me"],
      ["an unclaimed", null],
    ])("takes rewriting and deleting %s comment", (_name, authorId) => {
      const state = commented(authorId);
      const id = commentId(state);
      for (const command of [
        updateComment(id, "rewritten"),
        removeComment(id),
      ]) {
        expect(
          protectionAllows(
            state.doc,
            applied(state, command).doc,
            rules("comments")
          )
        ).toBe(true);
      }
    });

    it("takes carrying one's own comment onto other text", () => {
      const state = commented("me");
      expect(protectionAllows(state.doc, moved(state), rules("comments"))).toBe(
        true
      );
    });

    it("lets the owner take a thread down with the replies others left on it", () => {
      const state = commented("me");
      const id = commentId(state);
      const replied = applied(
        state,
        addCommentReply(id, { text: "reply", author: "Bo", authorId: "other" })
      );
      const removed = applied(replied, removeComment(id));
      expect(
        protectionAllows(replied.doc, removed.doc, rules("comments"))
      ).toBe(true);
    });
  });

  describe("over another author's reply", () => {
    const state = commented("me");
    const id = commentId(state);
    const replied = applied(
      state,
      addCommentReply(id, { text: "reply", author: "Bo", authorId: "other" })
    );
    const replyId = (() => {
      let found: string | null = null;
      replied.doc.descendants((node) => {
        if (node.type.name === "commentReference") {
          const replies: unknown = node.attrs.replies;
          if (Array.isArray(replies) && typeof replies[0]?.id === "string") {
            found = replies[0].id;
          }
        }
      });
      if (found === null) throw new Error("no reply");
      return found;
    })();

    it("refuses rewriting and deleting it, though the root is one's own", () => {
      for (const command of [
        updateCommentReply(id, replyId, "rewritten"),
        removeCommentReply(id, replyId),
      ]) {
        expect(
          protectionAllows(
            replied.doc,
            applied(replied, command).doc,
            rules("comments")
          )
        ).toBe(false);
      }
    });
  });
});

describe("a comment's identity", () => {
  it.each([
    ["another identity", { authorId: "me" }],
    ["another name", { author: "Me" }],
    ["no identity at all", { authorId: null }],
  ])("is never rewritten into %s", (_name, attrs) => {
    const state = commented("other");
    const after = rewritten(state, attrs);
    expect(protectionAllows(state.doc, after, rules("comments"))).toBe(false);
    expect(protectionAllows(state.doc, after, rules("none"))).toBe(false);
    expect(
      protectionAllows(state.doc, after, rules("comments", "me", "all"))
    ).toBe(false);
    expect(commentAdditionsBy(state.doc, after, "me")).toBe(false);
  });

  it("is not the owner's to rewrite either", () => {
    const state = commented("me");
    const after = rewritten(state, { author: "Someone else" });
    expect(protectionAllows(state.doc, after, rules("comments"))).toBe(false);
  });

  it("cannot be taken over by a rename in one round and a rewrite in the next", () => {
    const state = commented("other");
    const renamed = stateOf(rewritten(state, { authorId: "me", author: "Me" }));
    const edited = applied(
      renamed,
      updateComment(commentId(renamed), "rewritten")
    ).doc;
    // Judged from the renamed document the rewrite reads as the owner's own, so the rename is
    // where the takeover has to be refused, and it is refused however far back the judgement runs
    expect(protectionAllows(renamed.doc, edited, rules("comments"))).toBe(true);
    expect(protectionAllows(state.doc, renamed.doc, rules("comments"))).toBe(
      false
    );
    expect(protectionAllows(state.doc, edited, rules("comments"))).toBe(false);
  });
});

describe("commentAdditionsBy", () => {
  it("holds when every comment and reply that appeared carries the identity", () => {
    const before = opened();
    const mine = commented("me");
    expect(commentAdditionsBy(before.doc, mine.doc, "me")).toBe(true);
    const replied = applied(
      mine,
      addCommentReply(commentId(mine), {
        text: "r",
        author: "Me",
        authorId: "me",
      })
    );
    expect(commentAdditionsBy(mine.doc, replied.doc, "me")).toBe(true);
  });

  it("does not hold for a comment or reply claiming another identity, or none", () => {
    const before = opened();
    expect(commentAdditionsBy(before.doc, commented("other").doc, "me")).toBe(
      false
    );
    expect(commentAdditionsBy(before.doc, commented(null).doc, "me")).toBe(
      false
    );
    const mine = commented("me");
    const replied = applied(
      mine,
      addCommentReply(commentId(mine), {
        text: "r",
        author: "Bo",
        authorId: "other",
      })
    );
    expect(commentAdditionsBy(mine.doc, replied.doc, "me")).toBe(false);
  });
});
