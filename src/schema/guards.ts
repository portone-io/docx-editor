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
 *
 * A command decides nothing of its own about any of this. It takes one of the two shapes an edit
 * comes in - `guardedCommand`, which builds the whole edit and is refused whole, and
 * `openStretches`, which leaves the shut stretches out and applies to the rest - and both ask the
 * list below. A guard registered here therefore reaches every command by being registered, rather
 * than by each command being taught about it.
 *
 * `editsShut` (`./protectionState`) stays what it was, the view-level question of whether the body
 * is open at all, which is a question about the editor rather than about an edit.
 */

import type {
  Command,
  EditorState,
  Selection,
  Transaction,
} from "prosemirror-state";
import {
  type ChangeGuard,
  type EditGuard,
  type EditIntent,
  type StepGuard,
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

export type { EditGuard, EditGuardName, EditIntent } from "./editGuard";
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
const protectionGuard: ChangeGuard = {
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
 * What a command doing this to one stretch means to do.
 *
 * A stretch of no length holds nothing to mark or to put away, so whatever the command would do to
 * a stretch it is an insertion there.
 */
function intentOver(
  from: number,
  to: number,
  kind: "mark" | "replace"
): EditIntent {
  return from === to ? { kind: "insert", at: from } : { kind, from, to };
}

/** What a command doing this to whatever is selected means to do, one intent per selected stretch */
export function selectionIntents(
  selection: Selection,
  kind: "mark" | "replace"
): EditIntent[] {
  return selection.ranges.map((range) =>
    intentOver(range.$from.pos, range.$to.pos, kind)
  );
}

/**
 * Build, guard, dispatch: the one shape of a command that is refused whole.
 *
 * A structural edit has no smaller piece to fall back on - half a row cannot be deleted, and half a
 * block of cells cannot be merged into one - so the whole transaction is built, handed to the whole
 * guard list, and dispatched only if it comes back allowed. A character or a paragraph edit does
 * the opposite and leaves the shut stretches out (`openStretches`).
 *
 * The answer is the same whether or not `dispatch` was passed. The transaction is built before
 * either way, so asking the guard costs nothing more, and a command reporting one thing to a button
 * and doing another would be worse than the button being wrong.
 */
export function guardedCommand(
  build: (state: EditorState) => Transaction | null
): Command {
  return (state, dispatch) => {
    const tr = build(state);
    if (tr === null || !transactionAllowed(tr, state)) return false;
    dispatch?.(tr);
    return true;
  };
}

/**
 * The stretches the guards leave open, which is the one shape of a command that is trimmed.
 *
 * A shut stretch is left out rather than the whole edit refused: a guard turns down the whole
 * transaction, so asking for the shut stretch as well would leave the rest of the selection
 * unedited too. A selection the guards leave nothing of edits nothing, and the command reports that
 * of its own accord, which is the disabled state of the control that runs it.
 */
export function openStretches<S extends { from: number; to: number }>(
  state: EditorState,
  stretches: readonly S[],
  kind: "mark" | "replace"
): S[] {
  return stretches.filter(
    (stretch) => !editShut(state, intentOver(stretch.from, stretch.to, kind))
  );
}
