/**
 * The one edit the editor makes at the edge of a block-level content control: a control left
 * holding a single empty paragraph is removed whole (`spec/notes/contentControls.md`).
 *
 * The step is one covering the control from end to end, which is what `schema/locks` reads as the
 * deletion clause's question, so a control locked against deletion refuses it.
 */

import type { ResolvedPos } from "prosemirror-model";
import { type Command, Selection, TextSelection } from "prosemirror-state";
import { isBlockControl } from "../schema/controlAttrs";
import { guardedCommand } from "../schema/guards";

/**
 * The depth of the outermost control that holds nothing but the empty paragraph this caret stands
 * in. Null where the caret is somewhere else, or where a control still holds something.
 *
 * The walk passes through controls alone: a control may hold nothing but another control, and
 * removing the inner one alone would leave the outer one with nothing, which `block+` does not
 * admit. A cell, a row or a table standing in the way ends it, since the empty paragraph is not
 * all such a control holds.
 */
function soleEmptyControl($cursor: ResolvedPos): number | null {
  if ($cursor.parent.content.size > 0) return null;
  let found: number | null = null;
  for (let depth = $cursor.depth - 1; depth > 0; depth -= 1) {
    const node = $cursor.node(depth);
    if (node.childCount !== 1 || !isBlockControl(node)) break;
    found = depth;
  }
  return found;
}

export const removeEmptyBlockControl: Command = guardedCommand((state) => {
  const { selection } = state;
  const $cursor = selection instanceof TextSelection ? selection.$cursor : null;
  if (!$cursor) return null;
  const depth = soleEmptyControl($cursor);
  if (depth === null) return null;

  const from = $cursor.before(depth);
  const tr = state.tr.delete(from, $cursor.after(depth));
  // A control standing alone inside something that admits no empty content cannot be taken out
  if (!tr.docChanged) return null;
  return tr
    .setSelection(Selection.near(tr.doc.resolve(from), -1))
    .scrollIntoView();
});
