/**
 * Where the built-in comment composer stands: the stretch of text it was opened over, or nothing
 * when it is shut.
 *
 * The stretch is the point of holding it here rather than in the component that draws the form.
 * A comment is written about text the reader picked out, and writing it takes as long as it takes:
 * the caret moves on, the page is scrolled, a word elsewhere is fixed. The stretch is moved along
 * by every edit that happens meanwhile, so the comment lands on the text it was written for, and
 * the form goes away by itself once that text is gone or the document stops taking comments.
 *
 * The commands here are the editor's own. A consumer opening its own composer from its own toolbar
 * is a public surface this package has not settled yet, so nothing below is exported from
 * `./commands`.
 */

import type { Command, EditorState, Plugin } from "prosemirror-state";
import { canAddComment } from "../commands/comments/editing";
import type { CommentRange } from "../commands/comments/model";
import { panelPlugin } from "./panelState";

function isCommentRange(value: unknown): value is CommentRange {
  if (typeof value !== "object" || value === null) return false;
  const { from, to }: Partial<CommentRange> = value;
  return typeof from === "number" && typeof to === "number";
}

const composer = panelPlugin<CommentRange>({
  name: "docxEditorCommentComposer",
  isAnchor: isCommentRange,
  // The stretch moves with the text around it. Its two ends are biased inwards, so text typed
  // against either edge stays outside the comment and an edit that swallows the stretch leaves
  // the two ends together, which is the form closing
  onDocChange: (anchor, tr) => {
    const from = tr.mapping.map(anchor.from, 1);
    const to = tr.mapping.map(anchor.to, -1);
    return to > from ? { from, to } : null;
  },
  // Asked after every transaction, which is what catches the mode being switched to read-only:
  // that changes no text, so the mapping above never sees it. It is also what keeps the form from
  // ever standing over a stretch that takes no comment, whoever opened it
  closeWhen: (state, anchor) => !canAddComment(state, anchor),
});

/**
 * Opens the composer over the selected text. False where a comment could not go there anyway, so
 * the menu entry that runs it is drawn dead rather than opening a form nothing would come of. It
 * is the same question that entry is drawn from, asked of the same selection.
 */
export const openCommentComposer: Command = (state, dispatch) => {
  if (!canAddComment(state)) return false;
  const { from, to } = state.selection;
  return composer.open({ from, to })(state, dispatch);
};

/** Closes it. Called on Cancel, on the comment being written, and on the panel being put away */
export const closeCommentComposer: Command = composer.close;

export function isCommentComposerOpen(state: EditorState): boolean {
  return composer.anchor(state) !== null;
}

/** The stretch the open composer is writing about. Null when no composer stands */
export function commentComposerRange(state: EditorState): CommentRange | null {
  return composer.anchor(state);
}

export function commentComposer(): Plugin<CommentRange | null> {
  return composer.plugin;
}
