/**
 * What the editor does beside a content control the file wrote inside a paragraph with nothing in
 * it (`docx/wrappers`): Backspace and Delete take away what stands on the far side of it, as if the
 * control were not there (`spec/notes/contentControls.md`).
 *
 * Such a control draws nothing and takes no width, so a key that took it away would take a phrase
 * the file keeps with nothing on screen to show what went, and a key that merely passed the caret
 * over it would be a keystroke the user sees nothing at all come of. So the two are one keystroke:
 * the caret passes the whole run of them and the same key takes the character beyond. The control
 * itself goes when a stretch covering it is taken away, which is the deletion clause's question
 * (`schema/locks`).
 *
 * Between blocks the rule is `./blockControlEdits`'s, where passing the caret to another line is
 * something the user can see and the key stops there.
 */

import {
  chainCommands,
  joinBackward,
  joinForward,
  selectNodeBackward,
  selectNodeForward,
} from "prosemirror-commands";
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { isEmptyInlineControl } from "../schema/controlAttrs";
import { guardedCommand } from "../schema/guards";

type Direction = -1 | 1;

/** One stretch of the document, as a step covers it */
interface Reach {
  from: number;
  to: number;
}

/**
 * Where the caret lands once it has passed every such control standing on this side of it, and
 * null where none does.
 *
 * A whole run is passed at once: several of them in a row draw nothing between them either, so
 * stopping inside the run would be stopping nowhere the user can see.
 */
function beyondEmptyControls(
  doc: PMNode,
  pos: number,
  dir: Direction
): number | null {
  let at = pos;
  for (;;) {
    const $at = doc.resolve(at);
    const sibling = dir < 0 ? $at.nodeBefore : $at.nodeAfter;
    if (!sibling || !isEmptyInlineControl(sibling)) break;
    at += dir * sibling.nodeSize;
  }
  return at === pos ? null : at;
}

/** Where a caret standing in a textblock would pass such a run to, and null where none stands there */
function beyondRun(state: EditorState, dir: Direction): number | null {
  const $cursor =
    state.selection instanceof TextSelection ? state.selection.$cursor : null;
  return $cursor === null
    ? null
    : beyondEmptyControls(state.doc, $cursor.pos, dir);
}

/**
 * How much of the text one keystroke takes from this end of it.
 *
 * A surrogate pair is one character on the page and half of one is not text at all, so the pair
 * goes together.
 */
function charSize(text: string, dir: Direction): number {
  const point = text.codePointAt(dir < 0 ? Math.max(text.length - 2, 0) : 0);
  return point !== undefined && point > 0xffff ? 2 : 1;
}

/**
 * What the key takes from beyond the run: the character there, or the whole of a node that draws
 * as one thing. null where the block ends there and there is nothing beside it to take.
 */
function reachBeyond(
  doc: PMNode,
  beyond: number,
  dir: Direction
): Reach | null {
  const $beyond = doc.resolve(beyond);
  const sibling = dir < 0 ? $beyond.nodeBefore : $beyond.nodeAfter;
  if (!sibling) return null;
  const size = sibling.isText
    ? charSize(sibling.text ?? "", dir)
    : sibling.nodeSize;
  return dir < 0
    ? { from: beyond - size, to: beyond }
    : { from: beyond, to: beyond + size };
}

/**
 * What the key does where the run reaches the end of the block: whatever it would do with the
 * caret standing there, which is a join with the block beyond.
 *
 * The move the caret makes changes no content, so the steps built over it still stand on the
 * document the key was pressed over.
 */
function fromBlockEdge(
  state: EditorState,
  beyond: number,
  atEdge: Command
): Transaction | null {
  const moved = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, beyond))
  );
  let built: Transaction | null = null;
  atEdge(moved, (tr) => {
    built = tr;
  });
  return built;
}

/**
 * Passes the caret over the controls holding nothing that stand beside it and takes away what
 * stands beyond them.
 *
 * The caret itself does not move: what goes is on the far side of the run, so the controls stay
 * where they stood and the words either side of them close up as they would with no control there.
 */
function deleteBeyondEmptyControls(dir: Direction): Command {
  const atEdge =
    dir < 0
      ? chainCommands(joinBackward, selectNodeBackward)
      : chainCommands(joinForward, selectNodeForward);
  const edit = guardedCommand((state) => {
    const beyond = beyondRun(state, dir);
    if (beyond === null) return null;
    const reach = reachBeyond(state.doc, beyond, dir);
    if (reach === null) return fromBlockEdge(state, beyond, atEdge);
    return state.tr.delete(reach.from, reach.to).scrollIntoView();
  });
  return (state, dispatch, view) =>
    // With nothing beyond the run to take, the key still may not take the control away, so it is
    // reported handled and the document is left as it stood.
    edit(state, dispatch, view) || beyondRun(state, dir) !== null;
}

/** Backspace beside the controls holding nothing that stand before the caret */
export const deleteBeyondEmptyControlsBefore: Command =
  deleteBeyondEmptyControls(-1);

/** Delete beside the controls holding nothing that stand after the caret */
export const deleteBeyondEmptyControlsAfter: Command =
  deleteBeyondEmptyControls(1);
