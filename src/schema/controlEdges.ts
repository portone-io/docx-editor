/**
 * The edge of a block-level content control, which a keystroke's join may not carry blocks across
 * (`spec/notes/contentControls.md`).
 *
 * A control is not isolating (`./docxSchema`), so ProseMirror's base keymap will join across its
 * edge: the join a keystroke builds there may not carry blocks into a control or out of one, and
 * it is refused here rather than in the keymap, as one rule the whole guard list asks (`./guards`).
 *
 * Judging one step at a time is what confines the rule to the join. A cut and a paste, and the
 * drag that is the two in one transaction, are the user saying where the text goes; taking the
 * control away whole carries nothing across at all, and `./locks` answers that under the deletion
 * clause.
 */

import type { Node as PMNode, ResolvedPos } from "prosemirror-model";
import {
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
} from "prosemirror-transform";
import { isBlockControl } from "./controlAttrs";
import type { EditIntent, StepGuard } from "./editGuard";
import { controlLifted, historyReplay } from "./locks";

/**
 * Where each block control this position stands inside begins, outermost first. Empty when it
 * stands in none.
 *
 * A control's identity within one document is where it stands: `key` is not it, since a control
 * pasted from another keeps the key of the one it was copied from.
 */
function controlsAround($pos: ResolvedPos): number[] {
  const starts: number[] = [];
  for (let depth = 1; depth <= $pos.depth; depth += 1) {
    if (isBlockControl($pos.node(depth))) starts.push($pos.before(depth));
  }
  return starts;
}

function sameControls(a: readonly number[], b: readonly number[]): boolean {
  return a.length === b.length && a.every((start, index) => start === b[index]);
}

/**
 * Whether the two ends of the stretch stand inside different controls, which is the shape of a
 * replacement that would have to join across an edge to close itself.
 *
 * A stretch covering a control from the outside on both sides has the same controls at either end
 * - none of them the control it covers - so taking one away whole is not this.
 */
function crossesAnEdge(doc: PMNode, from: number, to: number): boolean {
  if (from === to) return false;
  return !sameControls(
    controlsAround(doc.resolve(from)),
    controlsAround(doc.resolve(to))
  );
}

/**
 * Whether the step puts the content it preserves inside a different set of controls.
 *
 * A `ReplaceAroundStep` hands back what stands between `gapFrom` and `gapTo` untouched and puts it
 * at `from + insert`, which is how ProseMirror both lifts a block out of its parent and wraps a
 * block into the one before it. The two spots stand in two documents, so the controls the gap came
 * from are carried over by the step's own map before they are compared: a control rewritten where
 * it stands (`setNodeMarkup`, which is how a lock is lifted) maps onto itself and reads the same
 * on both sides.
 */
function movesGapAcross(step: Step, before: PMNode, after: PMNode): boolean {
  if (!(step instanceof ReplaceAroundStep)) return false;
  const map = step.getMap();
  const was = controlsAround(before.resolve(step.gapFrom)).map((start) =>
    map.map(start, -1)
  );
  return !sameControls(
    was,
    controlsAround(after.resolve(step.from + step.insert))
  );
}

function stepAllowed(step: Step, before: PMNode, after: PMNode): boolean {
  if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) {
    return true;
  }
  return (
    !crossesAnEdge(before, step.from, step.to) &&
    !movesGapAcross(step, before, after)
  );
}

/**
 * The edges of the document's block controls, as `./guards` registers them.
 *
 * The history pass lifts it for the reason it lifts the locks: every step replayed is the reverse
 * of one that passed the guard when it was made, and a control the user deleted whole is put back
 * whole by the step that undoes it.
 */
export const controlEdgeGuard: StepGuard = {
  name: "controlEdge",
  liftedBy: [controlLifted, historyReplay],
  step: (step, before, after) => stepAllowed(step, before, after),
  shuts: (intent: EditIntent, state) =>
    intent.kind === "replace" &&
    crossesAnEdge(state.doc, intent.from, intent.to),
};
