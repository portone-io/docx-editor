// @vitest-environment jsdom
import {
  type Command,
  type EditorState,
  Plugin,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { rangeOfText } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { transactionAllowed } from "../../schema/locks";
import type { EditingProtection } from "../../schema/protection";
import {
  addComment,
  addCommentReply,
  canAddComment,
  canEditComment,
  removeComment,
  setCommentResolved,
  updateComment,
} from "../commands/commentCommands";
import { toggleBold } from "../commands/formattingCommands";
import { redo, undo } from "../commands/historyCommands";
import {
  createEditorState,
  createEditorView,
  type EditorStateOptions,
} from "../createEditor";
import { documentDefaults } from "../documentStyles";
import { setProtection } from "./documentProtection";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const BODY = `<w:p>${run("Alpha ")}${run("beta")}</w:p>`;

function opened(options: EditorStateOptions = {}): EditorState {
  return createEditorState(importDocx(makeDocx(BODY)).doc, options);
}

function selecting(state: EditorState, text: string): EditorState {
  const { from, to } = rangeOfText(state.doc, text);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

/** Runs the command through the state, which is where the guard stands, and says what came of it */
function attempt(
  state: EditorState,
  command: Command
): { answered: boolean; changed: boolean } {
  let after = state;
  const answered = command(state, (tr) => {
    after = after.apply(tr);
  });
  return { answered, changed: !after.doc.eq(state.doc) };
}

const REFUSED = { answered: false, changed: false };
const DONE = { answered: true, changed: true };

function commentId(state: EditorState): string {
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
}

/** A state under this protection holding a comment written under another identity */
function withOthersComment(
  protection: EditingProtection,
  editableComments: EditorStateOptions["editableComments"] = "own"
): EditorState {
  const written = selecting(opened(), "beta");
  let commented = written;
  addComment({ text: "note", author: "Bo", authorId: "other" })(
    written,
    (tr) => (commented = written.apply(tr))
  );
  return createEditorState(commented.doc, {
    protection,
    author: { id: "me", name: "Me" },
    editableComments,
  });
}

describe("the protection plugin", () => {
  it("refuses typing under readOnly and comments and takes it under none", () => {
    for (const protection of ["readOnly", "comments"] as const) {
      const state = opened({ protection });
      expect(state.apply(state.tr.insertText("x", 1)).doc.eq(state.doc)).toBe(
        true
      );
    }
    const open = opened({ protection: "none" });
    expect(open.apply(open.tr.insertText("x", 1)).doc.eq(open.doc)).toBe(false);
  });

  it("refuses a comment under readOnly and takes one under comments", () => {
    const comment = addComment({ text: "note", author: "Me", authorId: "me" });
    const reading = selecting(opened({ protection: "readOnly" }), "beta");
    expect(canAddComment(reading)).toBe(false);
    expect(attempt(reading, comment)).toEqual(REFUSED);

    const commenting = selecting(
      opened({ protection: "comments", author: { id: "me", name: "Me" } }),
      "beta"
    );
    expect(canAddComment(commenting)).toBe(true);
    expect(attempt(commenting, comment)).toEqual(DONE);
    expect(attempt(commenting, toggleBold)).toEqual(REFUSED);
  });

  it("keeps a reply, a resolution and the selection open on another author's comment", () => {
    const state = withOthersComment("comments");
    const id = commentId(state);
    expect(
      attempt(
        state,
        addCommentReply(id, { text: "r", author: "Me", authorId: "me" })
      )
    ).toEqual(DONE);
    expect(attempt(state, setCommentResolved(id, true))).toEqual(DONE);
  });

  it("shuts rewriting and deleting another author's comment under own, and opens them under all", () => {
    const own = withOthersComment("comments");
    const id = commentId(own);
    expect(canEditComment(own, id)).toBe(false);
    expect(attempt(own, updateComment(id, "rewritten"))).toEqual(REFUSED);
    expect(attempt(own, removeComment(id))).toEqual(REFUSED);

    const all = withOthersComment("comments", "all");
    expect(canEditComment(all, id)).toBe(true);
    expect(attempt(all, updateComment(id, "rewritten"))).toEqual(DONE);
    expect(attempt(all, removeComment(id))).toEqual(DONE);
  });

  it("answers no edit at all for a comment under readOnly", () => {
    const state = withOthersComment("readOnly", "all");
    expect(canEditComment(state, commentId(state))).toBe(false);
  });

  it("takes back one's own comment under comments and nothing under readOnly", () => {
    const commenting = selecting(
      opened({ protection: "comments", author: { id: "me", name: "Me" } }),
      "beta"
    );
    let after = commenting;
    addComment({ text: "note", author: "Me", authorId: "me" })(
      commenting,
      (tr) => (after = commenting.apply(tr))
    );
    const undone = attempt(after, undo);
    expect(undone).toEqual(DONE);

    let back = after;
    undo(after, (tr) => (back = after.apply(tr)));
    expect(attempt(back, redo)).toEqual(DONE);

    // A history holding an edit made before the protection shut: the replay is refused, and undo says so
    const edited = opened({ protection: "none" });
    const typed = edited.apply(edited.tr.insertText("x", 1));
    const shut = typed.apply(
      setProtection(typed.tr, {
        protection: "readOnly",
        author: null,
        editableComments: "own",
      })
    );
    expect(attempt(shut, undo)).toEqual(REFUSED);
  });
});

describe("switching the protection on an open state", () => {
  it("takes effect through the transaction and not through reconfigure", () => {
    const open = opened({ protection: "none" });
    const reconfigured = open.reconfigure({
      plugins: opened({ protection: "readOnly" }).plugins,
    });
    expect(
      reconfigured.apply(reconfigured.tr.insertText("x", 1)).doc.eq(open.doc)
    ).toBe(false);

    const shut = open.apply(
      setProtection(open.tr, {
        protection: "readOnly",
        author: null,
        editableComments: "own",
      })
    );
    expect(shut.apply(shut.tr.insertText("x", 1)).doc.eq(open.doc)).toBe(true);
  });
});

describe("the editor view under protection", () => {
  const mounts: HTMLElement[] = [];
  afterEach(() => {
    for (const mount of mounts) mount.remove();
    mounts.length = 0;
  });

  it.each([
    ["none", true],
    ["comments", false],
    ["readOnly", false],
  ] as const)("is editable under %s: %s", (protection, editable) => {
    const mount = document.body.appendChild(document.createElement("div"));
    mounts.push(mount);
    const state = opened({ protection });
    const view = createEditorView({
      mount,
      state,
      defaults: documentDefaults(state),
      onStateChange: () => {},
    });
    expect(view.editable).toBe(editable);
    view.destroy();
  });
});

/**
 * The plugins that put a document right after an edit - the derived table lines
 * (`table/gridBorders`) and the styles read into new paragraphs (`styledParagraphs`) - append
 * transactions of their own. An appended transaction the guard refuses is dropped without a word,
 * which would leave the document half corrected, so under `comments` nothing they append may be
 * refused.
 */
describe("the corrections the plugins append", () => {
  const cell = (text: string) =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
  const TABLE_BODY =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cell("Left")}${cell("Right")}</w:tr>` +
    "</w:tbl>";

  it("go through the guard under comments, table cells included", () => {
    const refused: Transaction[] = [];
    const watching = new Plugin({
      filterTransaction(tr, state) {
        if (!transactionAllowed(tr, state)) refused.push(tr);
        return true;
      },
    });
    const opened = createEditorState(
      importDocx(makeDocx(`${BODY}${TABLE_BODY}`)).doc,
      {
        protection: "comments",
        author: { id: "me", name: "Me" },
        consumerPlugins: [watching],
      }
    );

    let state = selecting(opened, "Left");
    const dispatch = (tr: Transaction) => {
      state = state.apply(tr);
    };
    expect(
      addComment({ text: "note", author: "Me", authorId: "me" })(
        state,
        dispatch
      )
    ).toBe(true);
    const id = commentId(state);
    for (const command of [
      addCommentReply(id, { text: "reply", author: "Me", authorId: "me" }),
      setCommentResolved(id, true),
      removeComment(id),
    ]) {
      expect(command(state, dispatch)).toBe(true);
    }

    expect(refused).toEqual([]);
    expect(state.doc.eq(opened.doc)).toBe(true);
  });
});
