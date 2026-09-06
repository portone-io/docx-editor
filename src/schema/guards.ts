/**
 * The one list an edit is judged against.
 *
 * A guard answers, for one rule, whether an edit would leave the document in a state the file
 * cannot be written back from. `transactionAllowed` runs the list over a built transaction, both
 * where an edit is dispatched (`editor/plugins/lockedContent`) and where a command is dry-run
 * (`editor/commands/canRunCommand`), and `editShut` runs it over the intent a command has before
 * it builds anything. Both ask the same list, which is what keeps a disabled control and a refused
 * edit saying the same thing.
 */

import type { Node as PMNode } from "prosemirror-model";
import type {
  EditorState,
  PluginKey,
  Selection,
  Transaction,
} from "prosemirror-state";
import {
  AddMarkStep,
  AddNodeMarkStep,
  AttrStep,
  RemoveMarkStep,
  RemoveNodeMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
} from "prosemirror-transform";
import { lockGuard } from "./locks";
import { bookmarkGuard, noteGuard } from "./preservedGuards";
import {
  isCommentNode,
  type ProtectionState,
  protectionAllows,
} from "./protection";
import { editsShut, protectionOf } from "./protectionState";

/**
 * What a command means to do, in as much of it as a guard can answer for before anything is built.
 *
 * The positions are the ones the command would work over, counted in the document as it stands.
 */
export type EditIntent =
  /** Something goes in at this spot, and nothing there goes away */
  | { kind: "insert"; at: number }
  /** What stands in this stretch is marked where it stands */
  | { kind: "mark"; from: number; to: number }
  /** What stands in this stretch goes away, whatever takes its place */
  | { kind: "replace"; from: number; to: number }
  /** The block at this spot is rewritten around its content, an alignment or an indent */
  | { kind: "block"; at: number };

export interface EditGuard {
  /** How this guard is named in the honesty test and in a refusal read by a developer */
  readonly name: string;
  /**
   * Whether the step may go through, judged over the document it was built against.
   * Omitted: every step passes.
   */
  step?(step: Step, before: PMNode, after: PMNode, state: EditorState): boolean;
  /** The judgement a rule needs both documents of the whole change at once for */
  change?(tr: Transaction, state: EditorState): boolean;
  /** Whether the intent is shut where it stands, which is what a command asks before it builds */
  shuts(intent: EditIntent, state: EditorState): boolean;
  /** The passes that lift this guard's step judgement. Omitted: none does */
  liftedBy?: readonly PluginKey<boolean>[];
}

/** Whether this stretch of the document holds a node the question answers for */
export function rangeHolds(
  doc: PMNode,
  from: number,
  to: number,
  holds: (node: PMNode) => boolean
): boolean {
  let found = false;
  doc.nodesBetween(from, to, (node) => {
    if (found) return false;
    if (holds(node)) found = true;
    return !found;
  });
  return found;
}

/**
 * Whether the step reaches a node the question answers for: puts one in, takes one out, or
 * rewrites the one where it stands.
 *
 * A rule about such nodes cannot have been broken by a change that reaches none of them, which is
 * what lets a guard settle the common transaction - typing, and nothing more - over the stretches
 * its own steps rewrote rather than over the whole document.
 *
 * A step of a kind this does not know - one a consumer brought - is answered as reaching one,
 * since what it rewrote is not known either. The whole-document judgement then has the say, and
 * an unknown step costs a comparison rather than a hole in the guard.
 */
export function stepReaches(
  step: Step,
  before: PMNode,
  after: PMNode,
  holds: (node: PMNode) => boolean
): boolean {
  if (
    step instanceof AttrStep ||
    step instanceof AddNodeMarkStep ||
    step instanceof RemoveNodeMarkStep
  ) {
    const node = before.nodeAt(step.pos);
    return node !== null && holds(node);
  }
  if (
    !(
      step instanceof ReplaceStep ||
      step instanceof ReplaceAroundStep ||
      step instanceof AddMarkStep ||
      step instanceof RemoveMarkStep
    )
  ) {
    return true;
  }
  let reached = false;
  step.getMap().forEach((oldStart, oldEnd, newStart, newEnd) => {
    reached ||=
      rangeHolds(before, oldStart, oldEnd, holds) ||
      rangeHolds(after, newStart, newEnd, holds);
  });
  return reached;
}

/** Whether any step of the transaction reaches a node the question answers for */
export function transactionReaches(
  tr: Transaction,
  holds: (node: PMNode) => boolean
): boolean {
  return tr.steps.some((step, index) =>
    stepReaches(step, tr.docs[index], tr.docs[index + 1] ?? tr.doc, holds)
  );
}

/**
 * Whether the transaction reaches a comment node, which is how a body, a reply and a resolution
 * change. Every comment lives in its three nodes, so a change that reaches none of them cannot
 * have changed a comment.
 */
export function transactionTouchesComments(tr: Transaction): boolean {
  return transactionReaches(tr, isCommentNode);
}

/**
 * Whether the protection lets this transaction through (`./protection`).
 *
 * The whole-document judgement is reached for only when a comment is touched at all
 * (`transactionTouchesComments`). A change that touches none is a body edit: through under `none`,
 * refused under `comments`, and nothing about ownership to ask.
 */
export function protectionAllowsTransaction(
  tr: Transaction,
  rules: ProtectionState
): boolean {
  switch (rules.protection) {
    case "readOnly":
      return false;
    case "none":
      return (
        !transactionTouchesComments(tr) ||
        protectionAllows(tr.before, tr.doc, rules)
      );
    case "comments":
      return (
        transactionTouchesComments(tr) &&
        protectionAllows(tr.before, tr.doc, rules)
      );
    default: {
      const unmodelled: never = rules.protection;
      return unmodelled;
    }
  }
}

/**
 * What the editor as a whole may receive, which is the reader's standing rather than anything the
 * document carries (`./protection`).
 *
 * It takes both documents at once so that it can tell a comment from everything else, and it has
 * no pass: a replayed edit is still an edit.
 */
const protectionGuard: EditGuard = {
  name: "protection",
  change: (tr, state) => protectionAllowsTransaction(tr, protectionOf(state)),
  shuts: (_intent, state) => editsShut(state),
};

export const EDIT_GUARDS: readonly EditGuard[] = [
  protectionGuard,
  lockGuard,
  bookmarkGuard,
  noteGuard,
];

type StepGuard = EditGuard & { step: NonNullable<EditGuard["step"]> };

function judgesSteps(guard: EditGuard): guard is StepGuard {
  return guard.step !== undefined;
}

/** Whether the transaction carries one of the passes through this guard */
function lifted(guard: EditGuard, tr: Transaction): boolean {
  return guard.liftedBy?.some((pass) => tr.getMeta(pass) === true) === true;
}

/**
 * Whether every guard would let this transaction through, decided and nothing else.
 *
 * The refusal a guard itself answers with carries a side effect - the composition it ends
 * (`editor/plugins/lockedContent`) - which a query about a button's state may not set off, so the
 * decision stands apart from it and every caller building an edit asks this rather than handing
 * the transaction to a state.
 *
 * The whole-change judgements come first and take no pass, since a pass lifts one guard's reading
 * of a step rather than another guard's reading of the change. What is left is judged step by
 * step, each step over the document it was built against.
 */
export function transactionAllowed(
  tr: Transaction,
  state: EditorState
): boolean {
  if (!tr.docChanged) return true;
  if (EDIT_GUARDS.some((guard) => guard.change?.(tr, state) === false)) {
    return false;
  }
  return EDIT_GUARDS.filter(judgesSteps)
    .filter((guard) => !lifted(guard, tr))
    .every((guard) =>
      // Each step counts positions in the document it was built against, which `docs` holds
      tr.steps.every((step, index) =>
        guard.step(
          step,
          tr.docs[index] ?? state.doc,
          tr.docs[index + 1] ?? tr.doc,
          state
        )
      )
    );
}

/**
 * Whether any guard shuts this intent where it stands.
 *
 * This is what a command asks before it reports that it applies: a command reporting true and then
 * being refused by the guard draws a live control that swallows the click.
 */
export function editShut(state: EditorState, intent: EditIntent): boolean {
  return EDIT_GUARDS.some((guard) => guard.shuts(intent, state));
}

/**
 * What a command doing this to whatever is selected means to do, one intent per selected stretch.
 *
 * A stretch of no length holds nothing to mark or to put away, so whatever the command would do to
 * a stretch it is an insertion there.
 */
export function selectionIntents(
  selection: Selection,
  kind: "mark" | "replace"
): EditIntent[] {
  return selection.ranges.map(
    (range): EditIntent =>
      range.$from.pos === range.$to.pos
        ? { kind: "insert", at: range.$from.pos }
        : { kind, from: range.$from.pos, to: range.$to.pos }
  );
}
