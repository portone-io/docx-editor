/**
 * What a guard is, and the reach primitives one is written with.
 *
 * This module names no guard and imports none, so a module that writes a guard (`./locks`,
 * `./preservedGuards`) and the list that registers them (`./guards`) can both read it without
 * pointing at each other. A guard reading its shape off `./guards` instead would put a cycle
 * there: `EDIT_GUARDS` is built while `./guards` is evaluated, so whenever the guard's own module
 * happened to be evaluated first the list would be built out of bindings that had not been
 * assigned yet, and every edit would then be judged against a list holding holes.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, PluginKey, Transaction } from "prosemirror-state";
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

/**
 * How a guard is named in the honesty test and in a refusal read by a developer.
 *
 * The names are written out because the registry is a closed list: `lockHonesty` checks each one
 * against the places it puts that guard to the test, and a place naming a guard that no longer
 * stands has to be a mistake the compiler catches rather than an annotation answering for nothing.
 */
export type EditGuardName =
  | "protection"
  | "lock"
  | "bookmark"
  | "note"
  | "section";

/** What every guard answers, whichever of the two judgements it is written as */
interface GuardCommon {
  readonly name: EditGuardName;
  /** Whether the intent is shut where it stands, which is what a command asks before it builds */
  shuts(intent: EditIntent, state: EditorState): boolean;
}

/**
 * A guard that judges one step at a time, over the document that step was built against.
 *
 * Only this form takes a pass. A pass lifts one guard's reading of a step, and a rule that has to
 * see both documents of the change at once has no single step for it to let through.
 */
export interface StepGuard extends GuardCommon {
  step(step: Step, before: PMNode, after: PMNode, state: EditorState): boolean;
  /** The passes that lift this guard's step judgement. Omitted: none does */
  liftedBy?: readonly PluginKey<boolean>[];
  change?: never;
}

/** A guard that needs both documents of the whole change at once */
export interface ChangeGuard extends GuardCommon {
  change(tr: Transaction, state: EditorState): boolean;
  step?: never;
  liftedBy?: never;
}

/**
 * One rule an edit is judged by: whether it would leave the document in a state the file cannot be
 * written back from.
 *
 * A guard is one judgement or the other, and answers `shuts` either way, so there is no writing
 * one that judges nothing and no hanging a pass on a judgement no pass can lift.
 */
export type EditGuard = StepGuard | ChangeGuard;

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
  // A mark step maps no position, so `getMap()` hands back `StepMap.empty` and the walk below
  // never runs: laying a mark over a node leaves the node itself standing, and reaches none.
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
