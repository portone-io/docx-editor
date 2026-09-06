/**
 * The one list an edit is judged against.
 *
 * A guard answers, for one rule, whether an edit would leave the document in a state the file
 * cannot be written back from. `transactionAllowed` runs the list over a built transaction, both
 * where an edit is dispatched (`editor/plugins/lockedContent`) and where a command is dry-run
 * (`editor/commands/canRunCommand`), and `editShut` runs it over the intent a command has before
 * it builds anything. Both ask the same list, which is what keeps a disabled control and a refused
 * edit saying the same thing.
 *
 * What a guard is written against - `EditGuard`, `EditIntent` and the reach primitives - stands in
 * `./editGuard` so that a module writing a guard need not read this one, and is handed on from
 * here so that a caller has one door to the whole seam.
 */

import type { EditorState, Selection, Transaction } from "prosemirror-state";
import {
  type EditGuard,
  type EditIntent,
  transactionReaches,
} from "./editGuard";
import { lockGuard } from "./locks";
import { bookmarkGuard, noteGuard } from "./preservedGuards";
import {
  isCommentNode,
  type ProtectionState,
  protectionAllows,
} from "./protection";
import { editsShut, protectionOf } from "./protectionState";

export type { EditGuard, EditIntent } from "./editGuard";
export { stepReaches } from "./editGuard";

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
