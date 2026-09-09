// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  type CommentEntry,
  commentedDocx,
  commentedRun,
  run,
} from "../../__testing__/comments";
import { decode } from "../../__testing__/docx";
import { rangeOfText, runCommand, select } from "../../__testing__/editing";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { docxSchema } from "../../schema";
import type {
  EditableComments,
  EditingProtection,
} from "../../schema/protection";
import { deleteRow } from "../../table";
import {
  addComment,
  addCommentReply,
  type CommentAuthor,
  type DocumentComment,
  documentComments,
  removeComment,
  selectComment,
} from "../commands/commentCommands";
import { redo, undo } from "../commands/historyCommands";
import { editorStateForSession } from "../createEditor";

interface Opened {
  state: EditorState;
  session: SessionStore;
}

interface Options {
  protection?: EditingProtection;
  author?: CommentAuthor;
  editableComments?: EditableComments;
}

function opened(
  body: string,
  comments: readonly CommentEntry[],
  options: Options = {}
): Opened {
  const imported = importDocx(commentedDocx(body, comments));
  return {
    state: editorStateForSession(imported, options),
    session: imported.session,
  };
}

const CHECK_THIS: readonly CommentEntry[] = [
  { id: "4", text: "Check this", author: "Ada" },
];

/** A comment inside a paragraph, with text of its own on either side of it */
const SURROUNDED = `<w:p>${run("start ")}${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`;

/** From inside the text before the comment to inside the text after it, markers and all */
function acrossTheComment(doc: PMNode): { from: number; to: number } {
  return {
    from: rangeOfText(doc, "start ").from + 3,
    to: rangeOfText(doc, " beta").from + 2,
  };
}

/** Where one of a comment's three nodes stands, which no public value reports for a marker */
function markerPos(doc: PMNode, type: string, id: string): number {
  let at = -1;
  doc.descendants((node, pos) => {
    if (at < 0 && node.type.name === type && node.attrs.id === id) at = pos;
    return at < 0;
  });
  if (at < 0) throw new Error(`no ${type} for comment ${id}`);
  return at;
}

/** The one comment the document holds, whatever became of the text it was written for */
function only(state: EditorState): DocumentComment {
  const comments = documentComments(state);
  expect(comments).toHaveLength(1);
  const comment = comments[0];
  if (comment === undefined) throw new Error("no comment in the document");
  return comment;
}

function referenceCount(doc: PMNode): number {
  let seen = 0;
  doc.descendants((node) => {
    if (node.type.name === "commentReference") seen += 1;
    return true;
  });
  return seen;
}

function exportedPart(
  state: EditorState,
  session: SessionStore,
  path: string
): string {
  return decode(unzipSync(exportDocx(state.doc, session))[path]);
}

/**
 * What `prosemirror-view` does for Backspace over a node no selection can hold
 * (`stopNativeHorizontalDelete`): the one path a range marker is deleted through, and one no
 * keymap command answers.
 */
function backspace(state: EditorState): EditorState {
  const before = state.selection.$head.nodeBefore;
  if (before === null) throw new Error("nothing stands before the caret");
  const size = before.isText ? 1 : before.nodeSize;
  const at = state.selection.$head.pos;
  return state.apply(state.tr.delete(at - size, at));
}

/** The same for Delete, which reads the node on the other side of the caret */
function forwardDelete(state: EditorState): EditorState {
  const after = state.selection.$head.nodeAfter;
  if (after === null) throw new Error("nothing stands after the caret");
  const size = after.isText ? 1 : after.nodeSize;
  const at = state.selection.$head.pos;
  return state.apply(state.tr.delete(at, at + size));
}

describe("a comment whose text an edit deletes", () => {
  it("stays anchored on the text typed over the stretch it marked", () => {
    const { state } = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`,
      CHECK_THIS
    );
    const marked = only(state);

    const typed = state.apply(
      state.tr.insertText("Gamma", marked.from, marked.to)
    );

    const comment = only(typed);
    expect(comment.anchored).toBe(true);
    expect(typed.doc.textContent).toBe("Gamma beta");
    expect(typed.doc.textBetween(comment.from, comment.to)).toBe("Gamma");
  });

  it("is left detached where a wider deletion swallowed its markers and reference", () => {
    const { state, session } = opened(SURROUNDED, CHECK_THIS);
    const replied = runCommand(
      state,
      addCommentReply("4", {
        text: "Seen",
        author: "Grace",
        date: "2026-08-23T00:00:00Z",
      })
    );
    const { from, to } = acrossTheComment(replied.doc);

    const deleted = replied.apply(replied.tr.delete(from, to));

    const comment = only(deleted);
    expect(comment.anchored).toBe(false);
    expect(comment.id).toBe("4");
    expect(comment.author).toBe("Ada");
    expect(comment.text).toBe("Check this");
    expect(comment.replies.map((reply) => reply.text)).toEqual(["Seen"]);
    expect(deleted.doc.textContent).toBe("staeta");
    // Where the deletion left off, so the thread is still read next to what took its text's place
    expect(comment.referencePos).toBe(from);

    const written = exportedPart(deleted, session, "word/comments.xml");
    expect(written).toContain("Check this");
    expect(written).toContain("Seen");
  });

  it("survives a cut, which is that deletion under the clipboard's name", () => {
    const { state } = opened(SURROUNDED, CHECK_THIS);
    const { from, to } = acrossTheComment(state.doc);
    const selected = select(state, from, to);

    // What `prosemirror-view` dispatches for a cut, meta and all
    const cut = selected.apply(
      selected.tr.deleteSelection().setMeta("uiEvent", "cut")
    );

    expect(only(cut).anchored).toBe(false);
    expect(cut.doc.textContent).toBe("staeta");
  });

  it("lands ahead of the text a paste puts in its place", () => {
    const { state } = opened(SURROUNDED, CHECK_THIS);
    const { from, to } = acrossTheComment(state.doc);
    const selected = select(state, from, to);

    const pasted = selected.apply(
      selected.tr.replaceSelection(
        new Slice(Fragment.from(docxSchema.text("pasted")), 0, 0)
      )
    );

    expect(pasted.doc.textContent).toBe("stapastedeta");
    const comment = only(pasted);
    expect(comment.anchored).toBe(false);
    // The reference leans on the text that stood here rather than on the text that arrived
    expect(comment.referencePos).toBe(from);
  });

  it("moves into a neighbouring paragraph when its own paragraph goes", () => {
    const { state } = opened(
      `<w:p>${commentedRun("4", "Alpha")}</w:p><w:p>${run("Second")}</w:p>`,
      CHECK_THIS
    );
    const paragraph = state.doc.child(0).nodeSize;

    const deleted = state.apply(state.tr.delete(0, paragraph));

    const comment = only(deleted);
    expect(comment.anchored).toBe(false);
    expect(deleted.doc.childCount).toBe(1);
    expect(deleted.doc.resolve(comment.referencePos).parent.textContent).toBe(
      "Second"
    );
  });

  it("keeps two comments swept away by one edit, in the order they stood in", () => {
    const { state } = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" and ")}` +
        `${commentedRun("5", "Beta")}${run(" end")}</w:p>`,
      [
        { id: "4", text: "First note", author: "Ada" },
        { id: "5", text: "Second note", author: "Ada" },
      ]
    );
    const end = state.doc.child(0).nodeSize - 5;

    const deleted = state.apply(state.tr.delete(1, end));

    expect(
      documentComments(deleted).map(({ id, text, anchored }) => ({
        id,
        text,
        anchored,
      }))
    ).toEqual([
      { id: "4", text: "First note", anchored: false },
      { id: "5", text: "Second note", anchored: false },
    ]);
  });

  it("is not put back where one of its own copies is still standing", () => {
    const { state } = opened(
      `<w:p>${commentedRun("4", "Alpha")}</w:p>` +
        `<w:p><w:r><w:commentReference w:id="4"/></w:r>${run("Second")}</w:p>`,
      CHECK_THIS
    );
    expect(referenceCount(state.doc)).toBe(2);

    const deleted = state.apply(
      state.tr.delete(0, state.doc.child(0).nodeSize)
    );

    expect(referenceCount(deleted.doc)).toBe(1);
    expect(only(deleted).anchored).toBe(false);
  });

  it("keeps the run mark, so the style Word draws the mark with is written again", () => {
    const { state, session } = opened(
      `<w:p>${run("start ")}<w:commentRangeStart w:id="4"/>${run("Alpha")}` +
        '<w:commentRangeEnd w:id="4"/><w:r><w:rPr>' +
        '<w:rStyle w:val="CommentReference"/></w:rPr>' +
        `<w:commentReference w:id="4"/></w:r>${run(" beta")}</w:p>`,
      CHECK_THIS
    );
    const { from, to } = acrossTheComment(state.doc);

    const deleted = state.apply(state.tr.delete(from, to));

    expect(only(deleted).anchored).toBe(false);
    expect(exportedPart(deleted, session, "word/document.xml")).toContain(
      '<w:rStyle w:val="CommentReference"/>'
    );
  });

  it("survives being exported and opened again", () => {
    const { state, session } = opened(SURROUNDED, CHECK_THIS);
    const { from, to } = acrossTheComment(state.doc);
    const deleted = state.apply(state.tr.delete(from, to));

    const reopened = importDocx(exportDocx(deleted.doc, session));

    const comment = only(editorStateForSession(reopened));
    expect(comment.anchored).toBe(false);
    expect(comment.text).toBe("Check this");
    expect(comment.author).toBe("Ada");
  });

  it("takes the caret to where it stands when `selectComment` asks for it", () => {
    const { state } = opened(SURROUNDED, CHECK_THIS);
    const { from, to } = acrossTheComment(state.doc);
    const deleted = state.apply(state.tr.delete(from, to));

    const selected = runCommand(deleted, selectComment("4"));

    expect(selected.selection.empty).toBe(true);
    expect(selected.selection.from).toBe(only(deleted).referencePos);
  });

  it("is put back although the comment is another author's", () => {
    const { state } = opened(`<w:p>${run("start Alpha beta")}</w:p>`, [], {
      author: { id: "grace", name: "Grace" },
      editableComments: "own",
    });
    const commented = runCommand(
      state,
      addComment(
        {
          text: "Ada's note",
          author: "Ada",
          authorId: "ada",
          date: "2026-08-22T01:02:03Z",
        },
        rangeOfText(state.doc, "Alpha")
      )
    );
    const { from, to } = acrossTheComment(commented.doc);

    const deleted = commented.apply(commented.tr.delete(from, to));

    const comment = only(deleted);
    expect(comment.authorId).toBe("ada");
    expect(comment.anchored).toBe(false);
  });
});

/** A cell a content control locks, which no insertion may reach */
function shutCell(text: string): string {
  return (
    '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '<w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
    `<w:sdtContent><w:tc><w:p>${run(text)}</w:p></w:tc></w:sdtContent></w:sdt>`
  );
}

describe("a home the guards would refuse", () => {
  it("is passed over for the nearest one they leave open", () => {
    const { state } = opened(
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:p>${run("open")}</w:p></w:tc>${shutCell("shut")}</w:tr>` +
        `</w:tbl><w:p>${commentedRun("4", "Alpha")}</w:p>`,
      CHECK_THIS
    );

    const deleted = state.apply(
      state.tr.delete(state.doc.child(0).nodeSize, state.doc.content.size)
    );

    const comment = only(deleted);
    expect(comment.anchored).toBe(false);
    // The cell the deletion left the caret nearest is locked, so the open one takes the thread
    expect(deleted.doc.resolve(comment.referencePos).parent.textContent).toBe(
      "open"
    );
  });

  it("loses the comment where every textblock is one, which is the plugin's one limit", () => {
    const { state } = opened(
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        `<w:tr>${shutCell("shut")}</w:tr></w:tbl>` +
        `<w:p>${commentedRun("4", "Alpha")}</w:p>`,
      CHECK_THIS
    );

    const deleted = state.apply(
      state.tr.delete(state.doc.child(0).nodeSize, state.doc.content.size)
    );

    expect(documentComments(deleted)).toEqual([]);
  });

  it("keeps a comment through a row deleted from under it", () => {
    const { state } = opened(
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:p>${run("first")}</w:p></w:tc></w:tr>` +
        `<w:tr><w:tc><w:p>${commentedRun("4", "Alpha")}</w:p></w:tc></w:tr>` +
        "</w:tbl>",
      CHECK_THIS
    );
    const inside = rangeOfText(state.doc, "Alpha").from;

    const deleted = runCommand(select(state, inside), deleteRow);

    const comment = only(deleted);
    expect(comment.anchored).toBe(false);
    expect(deleted.doc.textContent).toBe("first");
  });
});

describe("a range marker deleted on its own", () => {
  it("goes back, so the comment keeps the text that never moved", () => {
    const { state, session } = opened(SURROUNDED, CHECK_THIS);
    const end = markerPos(state.doc, "commentEnd", "4");

    const deleted = state.apply(state.tr.delete(end, end + 1));

    expect(only(deleted).anchored).toBe(true);
    expect(deleted.doc.eq(state.doc)).toBe(true);
    const document = exportedPart(deleted, session, "word/document.xml");
    expect(document).toContain('<w:commentRangeStart w:id="4"/>');
    expect(document).toContain('<w:commentRangeEnd w:id="4"/>');
  });

  it("keeps the comment on what is left when the tail goes with it", () => {
    const { state, session } = opened(SURROUNDED, CHECK_THIS);
    const reference = markerPos(state.doc, "commentReference", "4");
    const from = rangeOfText(state.doc, "Alpha").from + 2;

    const deleted = state.apply(state.tr.delete(from, reference + 1));

    const comment = only(deleted);
    expect(comment.anchored).toBe(true);
    expect(deleted.doc.textBetween(comment.from, comment.to)).toBe("Al");
    const document = exportedPart(deleted, session, "word/document.xml");
    expect(document).toContain('<w:commentRangeStart w:id="4"/>');
    expect(document).toContain('<w:commentRangeEnd w:id="4"/>');
  });

  it("lets Backspace through rather than standing in its way", () => {
    const { state } = opened(
      `<w:p>${run("Alpha ")}${commentedRun("4", "beta")}</w:p>`,
      CHECK_THIS
    );
    const caret = select(state, state.doc.child(0).nodeSize - 1);

    // The press deletes the reference, which goes straight back; the caret carries on past the
    // markers rather than facing the same one again
    const stepped = backspace(caret);
    expect(stepped.doc.eq(state.doc)).toBe(true);
    expect(only(stepped).anchored).toBe(true);

    const shortened = backspace(stepped);
    expect(shortened.doc.textContent).toBe("Alpha bet");
    expect(only(shortened).anchored).toBe(true);
  });

  it("lets Delete through the opening marker the same way", () => {
    const { state } = opened(
      `<w:p>${run("Alpha ")}${commentedRun("4", "beta")}</w:p>`,
      CHECK_THIS
    );
    const caret = select(state, markerPos(state.doc, "commentStart", "4"));

    const stepped = forwardDelete(caret);
    expect(stepped.doc.eq(state.doc)).toBe(true);

    const shortened = forwardDelete(stepped);
    expect(shortened.doc.textContent).toBe("Alpha eta");
    expect(only(shortened).anchored).toBe(true);
  });
});

describe("a comment taken down on purpose", () => {
  it("is not put back by the plugin that keeps a deleted one", () => {
    const { state, session } = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`,
      CHECK_THIS
    );

    const removed = runCommand(state, removeComment("4"));

    expect(documentComments(removed)).toEqual([]);
    expect(referenceCount(removed.doc)).toBe(0);
    expect(exportedPart(removed, session, "word/comments.xml")).not.toContain(
      "Check this"
    );
  });

  it("stays gone when the removal is redone", () => {
    const { state } = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`,
      CHECK_THIS
    );

    const removed = runCommand(state, removeComment("4"));
    const undone = runCommand(removed, undo);
    const redone = runCommand(undone, redo);

    expect(only(undone).anchored).toBe(true);
    expect(documentComments(redone)).toEqual([]);
  });
});

describe("taking a deletion back and making it again", () => {
  it("restores the anchored comment, then detaches it once, with no second reference", () => {
    const { state } = opened(SURROUNDED, CHECK_THIS);
    const { from, to } = acrossTheComment(state.doc);

    const deleted = state.apply(state.tr.delete(from, to));
    const undone = runCommand(deleted, undo);
    const redone = runCommand(undone, redo);

    expect(only(deleted).anchored).toBe(false);
    expect(only(undone).anchored).toBe(true);
    expect(undone.doc.textContent).toBe("start Alpha beta");
    expect(only(redone).anchored).toBe(false);
    expect(referenceCount(redone.doc)).toBe(1);
  });
});

describe("where the body takes no edit at all", () => {
  it.each(["comments", "readOnly"] as const)(
    "does nothing under %s, because the deletion it would answer is refused",
    (protection) => {
      const { state } = opened(SURROUNDED, CHECK_THIS, { protection });
      const { from, to } = acrossTheComment(state.doc);

      const attempted = state.apply(state.tr.delete(from, to));

      expect(attempted.doc.eq(state.doc)).toBe(true);
      expect(only(attempted).anchored).toBe(true);
    }
  );
});
