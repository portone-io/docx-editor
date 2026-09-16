/**
 * What the editor does at the edge of a block-level content control, beyond the join the edge
 * itself refuses (`schema/controlEdges`): a control left holding a single empty paragraph is
 * removed whole, and a control holding nothing is passed over rather than taken away, the caret's
 * own empty paragraph going as it would anywhere else (`spec/notes/contentControls.md`).
 *
 * The removal is one step covering the control from end to end, which is what `schema/locks` reads
 * as the deletion clause's question, so a control locked against deletion refuses it.
 */

import type { Node as PMNode, ResolvedPos } from "prosemirror-model";
import { type Command, Selection, TextSelection } from "prosemirror-state";
import { docxSchema } from "../schema";
import { isBlockControl, isEmptyControl } from "../schema/controlAttrs";
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

/** Which way a key reaches out of the block the caret stands in */
type Direction = -1 | 1;

/** The run of controls holding nothing that stands next to the caret's own block */
interface EmptyRun {
  /** Where the caret goes to stand on the far side of the run, null where nothing stands there */
  beyond: number | null;
}

/** A paragraph with nothing in it, which a key beside a control takes away rather than keeps */
function isEmptyParagraph(node: PMNode): boolean {
  return node.type === docxSchema.nodes.paragraph && node.content.size === 0;
}

/** The caret where it stands at the edge the key reaches out of, and null anywhere else */
function caretAtEdge(selection: Selection, dir: Direction): ResolvedPos | null {
  const $cursor = selection instanceof TextSelection ? selection.$cursor : null;
  if (!$cursor || $cursor.depth === 0) return null;
  const at = dir < 0 ? 0 : $cursor.parent.content.size;
  return $cursor.parentOffset === at ? $cursor : null;
}

/**
 * The controls holding nothing that stand between the caret's block and the block beyond them, and
 * null where the next sibling is anything else.
 *
 * A whole run is passed at once: several such controls in a row draw nothing between them either,
 * so stopping inside the run would be stopping nowhere the user can see.
 */
function emptyRunBeside(selection: Selection, dir: Direction): EmptyRun | null {
  const $cursor = caretAtEdge(selection, dir);
  if (!$cursor) return null;
  const depth = $cursor.depth;
  const container = $cursor.node(depth - 1);
  const first = $cursor.index(depth - 1) + (dir < 0 ? -1 : 1);
  let pos = dir < 0 ? $cursor.before(depth) : $cursor.after(depth);
  let index = first;
  let sibling: PMNode | null | undefined = container.maybeChild(index);
  while (sibling && isEmptyControl(sibling)) {
    pos += dir * sibling.nodeSize;
    index += dir;
    sibling = container.maybeChild(index);
  }
  if (index === first) return null;
  return { beyond: sibling ? pos : null };
}

/**
 * Passes the caret over the controls holding nothing that stand beside it, taking away nothing but
 * an empty paragraph of the caret's own.
 *
 * Such a control draws nothing on the page, so a key that took it away would take a clause the
 * file keeps with nothing on screen to show what went: the control is the slot a server re-renders
 * into (ADR-024). It goes only when something covers it - a selection of the node itself or a
 * stretch running over it - which is the deletion clause's question as it is anywhere else
 * (`schema/locks`).
 *
 * The blank line the caret stands on is not the control, so it goes as it would anywhere else: the
 * step covers that paragraph alone, never the controls the caret then passes over, and a paragraph
 * that cannot be taken out of what holds it leaves the caret to move by itself.
 */
function skipEmptyControls(dir: Direction): Command {
  const move = guardedCommand((state) => {
    const beyond = emptyRunBeside(state.selection, dir)?.beyond;
    if (beyond === null || beyond === undefined) return null;
    const $cursor =
      state.selection instanceof TextSelection ? state.selection.$cursor : null;
    if ($cursor && isEmptyParagraph($cursor.parent)) {
      const tr = state.tr.delete($cursor.before(), $cursor.after());
      if (tr.docChanged) {
        const at = tr.doc.resolve(tr.mapping.map(beyond));
        return tr.setSelection(Selection.near(at, dir)).scrollIntoView();
      }
    }
    return state.tr
      .setSelection(Selection.near(state.doc.resolve(beyond), dir))
      .scrollIntoView();
  });
  return (state, dispatch, view) =>
    // With nothing beyond the run there is nowhere to pass to, and the key still may not take the
    // control away, so it is reported handled and the document is left as it stood.
    move(state, dispatch, view) ||
    emptyRunBeside(state.selection, dir) !== null;
}

/** Backspace passing over the controls holding nothing that stand before the caret */
export const skipEmptyControlBefore: Command = skipEmptyControls(-1);

/** Delete passing over the controls holding nothing that stand after the caret */
export const skipEmptyControlAfter: Command = skipEmptyControls(1);
