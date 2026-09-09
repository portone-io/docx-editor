// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { commentedDocx, commentedRun, run } from "../../__testing__/comments";
import { runCommand } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { setCommentResolved } from "../commands/commentCommands";
import { editorStateForSession } from "../createEditor";
import { commentProjection } from "./commentDecorations";

function opened(body: string): EditorState {
  return editorStateForSession(
    importDocx(commentedDocx(body, [{ id: "4", text: "Check this" }]))
  );
}

function drawnRanges(state: EditorState) {
  return commentProjection
    .read(state)
    .decorations.find()
    .map(({ from, to }) => ({ from, to }));
}

function onlyComment(state: EditorState) {
  const comment = commentProjection.read(state).comments[0];
  if (comment === undefined) throw new Error("no comment in the document");
  return comment;
}

describe("what the comment projection makes of a document's anchors", () => {
  it("reads a comment over a stretch of text as anchored and draws its range", () => {
    const state = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`
    );
    const comment = onlyComment(state);

    expect(comment.anchored).toBe(true);
    expect(comment.from).toBeLessThan(comment.to);
    expect(drawnRanges(state)).toEqual([
      { from: comment.from, to: comment.to },
    ]);
  });

  it("reads markers standing next to each other as anchored, with nothing to draw", () => {
    const state = opened(
      '<w:p><w:commentRangeStart w:id="4"/><w:commentRangeEnd w:id="4"/>' +
        `<w:r><w:commentReference w:id="4"/></w:r>${run("Alpha")}</w:p>`
    );
    const comment = onlyComment(state);

    // The stretch is empty rather than missing: the comment still marks the spot it was written at
    expect(comment.anchored).toBe(true);
    expect(comment.from).toBe(comment.to);
    expect(drawnRanges(state)).toEqual([]);
  });

  it("reads a comment missing a range marker as detached", () => {
    const state = opened(
      `<w:p><w:commentRangeStart w:id="4"/>${run("Alpha")}` +
        '<w:r><w:commentReference w:id="4"/></w:r></w:p>'
    );

    expect(onlyComment(state).anchored).toBe(false);
    expect(drawnRanges(state)).toEqual([]);
  });

  it("stops drawing a comment's range once it is resolved, and draws it again on reopening", () => {
    const state = opened(
      `<w:p>${commentedRun("4", "Alpha")}${run(" beta")}</w:p>`
    );
    const range = drawnRanges(state);
    expect(range).toHaveLength(1);

    const resolved = runCommand(state, setCommentResolved("4", true));
    expect(onlyComment(resolved).anchored).toBe(true);
    expect(drawnRanges(resolved)).toEqual([]);

    const reopened = runCommand(resolved, setCommentResolved("4", false));
    expect(drawnRanges(reopened)).toEqual(range);
  });
});
