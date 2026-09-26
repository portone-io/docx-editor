/**
 * What an application is told when the editor turns an edit down, so that it can say why.
 *
 * The guards (`schema/guards`) decide; this only reads a refused transaction back into terms a
 * host can act on: which rule refused it, what the edit would have done, and, for a lock, which
 * controls it reached, named by what the template wrote on them.
 *
 * The controls are read off the stretches the refused steps rewrite, which is the same reach the
 * lock guard judges by, but it is a report rather than a second judgement: every control carrying a
 * lock that the edit met is named, whichever of them the guard found first.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
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
import { childElement, parseProps, propsChild } from "../ooxml/props";
import { localPart } from "../ooxml/xml";
import {
  type ControlFacts,
  controlAttrsOf,
  controlFactsOf,
  isEmptyBlockControl,
  isEmptyInlineControl,
  OWN_CONTROL_ATTRS,
} from "../schema/controlAttrs";
import { transactionRefusal } from "../schema/guards";
import { controlSpans, emptyControlsIn } from "../schema/locks";

/**
 * The rule that turned an edit down.
 *
 * - `protection`: the editor's `mode` takes no such edit, such as a body edit in a comment editor.
 * - `lock`: the edit reaches a content control that the document locked, or a group control.
 * - `controlEdge`: a keystroke would join text across the edge of a content control around blocks.
 * - `preserved`: the edit would take away or reorder content the editor keeps without editing it,
 *   such as a bookmark's ends or the pieces of a field.
 * - `section`: the edit would take a section break away.
 *
 * It is written out rather than read off the guard list, so that a guard added there has to be
 * named here before it can be reported.
 */
export type EditRefusalReason =
  | "protection"
  | "lock"
  | "controlEdge"
  | "preserved"
  | "section";

/**
 * What the refused edit would have done: put content in, take it away, both at once, or change
 * formatting or structure while leaving the content where it stands.
 */
export type EditRefusalAction = "insert" | "delete" | "replace" | "format";

/** The `w:lock` value a control carries (§17.18.49), `unlocked` when it carries none */
export type ControlLock =
  | "unlocked"
  | "sdtLocked"
  | "contentLocked"
  | "sdtContentLocked";

/** What a content control stands around: text in a paragraph, whole blocks, a cell, or a row */
export type ControlLevel = "inline" | "block" | "cell" | "row";

/** A content control that refuses edits to its contents or its own removal */
export interface LockedControl {
  /** The control's `w:tag`, which is how a template names it for software */
  tag: string | null;
  /** The control's `w:alias`, which is the name Word shows a reader */
  alias: string | null;
  /** The control's `w:id` */
  id: number | null;
  /** What its `w:lock` says about editing and removing it */
  lock: ControlLock;
  /** Whether it is a group control, which shuts its contents without a lock */
  group: boolean;
  /** What it stands around */
  level: ControlLevel;
  /** Where the control starts in the editor's document */
  pos: number;
}

/** An edit the editor turned down, as `onEditRefused` is handed it */
export interface EditRefusal {
  /** The rule that refused it */
  reason: EditRefusalReason;
  /** What it would have done */
  action: EditRefusalAction;
  /** Where the refused edit starts in the editor's document */
  pos: number;
  /** The locked controls the edit reached, innermost first. Empty unless `reason` is `lock` */
  controls: readonly LockedControl[];
}

/** The stretch one step rewrites, in the document it was built against */
interface Reach {
  from: number;
  to: number;
}

function reachesOf(step: Step, doc: PMNode): Reach[] {
  if (step instanceof ReplaceAroundStep) {
    return [
      { from: step.from, to: step.gapFrom },
      { from: step.gapTo, to: step.to },
    ];
  }
  if (
    step instanceof ReplaceStep ||
    step instanceof AddMarkStep ||
    step instanceof RemoveMarkStep
  ) {
    return [{ from: step.from, to: step.to }];
  }
  if (
    step instanceof AttrStep ||
    step instanceof AddNodeMarkStep ||
    step instanceof RemoveNodeMarkStep
  ) {
    const node = doc.nodeAt(step.pos);
    return node ? [{ from: step.pos, to: step.pos + node.nodeSize }] : [];
  }
  return [];
}

function actionOf(tr: Transaction): EditRefusalAction {
  const replaced = tr.steps.filter((step) => step instanceof ReplaceStep);
  const takes = replaced.some((step) => step.from < step.to);
  const puts = replaced.some((step) => step.slice.size > 0);
  if (takes && puts) return "replace";
  if (takes) return "delete";
  if (puts) return "insert";
  return "format";
}

function lockOf(facts: ControlFacts): ControlLock {
  if (facts.contentsLocked && facts.deletionLocked) return "sdtContentLocked";
  if (facts.contentsLocked) return "contentLocked";
  if (facts.deletionLocked) return "sdtLocked";
  return "unlocked";
}

function refuses(facts: ControlFacts): boolean {
  return facts.contentsLocked || facts.deletionLocked || facts.group;
}

/** The `w:val` of one child of the control's `w:sdtPr`, read off the opening XML it carries */
function sdtPrValue(prefix: string, name: string): string | null {
  const sdt = parseProps(`${prefix}</w:sdt>`);
  const child = sdt && propsChild(sdt.children, "sdtPr");
  const sdtPr = child ? parseProps(child.xml) : null;
  const element = sdtPr ? childElement(sdtPr, name) : null;
  const val = element?.attrs.find(([attr]) => localPart(attr) === "val");
  return val ? val[1] : null;
}

function idOf(prefix: string): number | null {
  const value = sdtPrValue(prefix, "id");
  const id = value === null ? Number.NaN : Number.parseInt(value, 10);
  return Number.isNaN(id) ? null : id;
}

function lockedControl(
  facts: ControlFacts,
  level: ControlLevel,
  pos: number
): LockedControl {
  const prefix = facts.prefix ?? "";
  return {
    tag: sdtPrValue(prefix, "tag"),
    alias: sdtPrValue(prefix, "alias"),
    id: idOf(prefix),
    lock: lockOf(facts),
    group: facts.group,
    level,
    pos,
  };
}

function containerLevel(node: PMNode): ControlLevel {
  if (isEmptyInlineControl(node)) return "inline";
  if (isEmptyBlockControl(node)) return "block";
  const role: unknown = node.type.spec.tableRole;
  if (role === "row") return "row";
  if (role === "cell") return "cell";
  return "block";
}

/** A locked control found by the walk, with how deep it stands so the innermost can lead */
interface Found {
  control: LockedControl;
  depth: number;
}

function containerFound(
  node: PMNode,
  pos: number,
  depth: number
): Found | null {
  const names = controlAttrsOf(node);
  if (!names) return null;
  const facts = controlFactsOf(names, node.attrs);
  if (facts.prefix === null || !refuses(facts)) return null;
  return {
    control: lockedControl(facts, containerLevel(node), pos),
    depth,
  };
}

function markFacts(mark: Mark): ControlFacts {
  return controlFactsOf(OWN_CONTROL_ATTRS, mark.attrs);
}

/** Whether this stretch reaches into a span, an insertion counting only strictly inside it */
function meets(reach: Reach, span: Reach): boolean {
  return reach.from === reach.to
    ? span.from < reach.from && reach.from < span.to
    : reach.from < span.to && reach.to > span.from;
}

/** The locked inline controls of one textblock that this stretch reaches */
function inlineFound(
  block: PMNode,
  start: number,
  depth: number,
  reach: Reach
): Found[] {
  const textblock = { node: block, start };
  const marked = controlSpans(textblock)
    .filter((span) => meets(reach, span))
    .filter((span) => refuses(markFacts(span.mark)))
    .map((span) => ({
      control: lockedControl(markFacts(span.mark), "inline", span.from),
      depth: depth + 1 + Number(span.mark.attrs.depth ?? 0),
    }));
  const empty = emptyControlsIn(textblock)
    .filter((control) => meets(reach, control.span))
    .map((control) => containerFound(control.node, control.span.from, depth))
    .filter((found) => found !== null);
  return [...marked, ...empty];
}

/** Every locked control this stretch stands inside or reaches into */
function controlsReached(doc: PMNode, reach: Reach): Found[] {
  const found: Found[] = [];
  for (const at of [reach.from, reach.to]) {
    const $pos = doc.resolve(at);
    for (let depth = 1; depth <= $pos.depth; depth += 1) {
      const container = containerFound(
        $pos.node(depth),
        $pos.before(depth),
        depth
      );
      if (container) found.push(container);
    }
    if ($pos.parent.isTextblock) {
      found.push(...inlineFound($pos.parent, $pos.start(), $pos.depth, reach));
    }
  }
  doc.nodesBetween(reach.from, reach.to, (node, pos) => {
    const depth = doc.resolve(pos).depth + 1;
    const container = containerFound(node, pos, depth);
    if (container) found.push(container);
    if (!node.isTextblock) return true;
    found.push(...inlineFound(node, pos + 1, depth, reach));
    return false;
  });
  return found;
}

function sameControl(a: LockedControl, b: LockedControl): boolean {
  return a.pos === b.pos && a.level === b.level && a.id === b.id;
}

/** The locked controls the transaction's steps reach, innermost first, placed in the state's document */
function controlsOf(tr: Transaction, state: EditorState): LockedControl[] {
  const found = tr.steps.flatMap((step, index) => {
    const doc = tr.docs[index] ?? state.doc;
    const back = tr.mapping.slice(0, index).invert();
    return reachesOf(step, doc)
      .flatMap((reach) => controlsReached(doc, reach))
      .map(({ control, depth }) => ({
        control: { ...control, pos: back.map(control.pos, 1) },
        depth,
      }));
  });
  const unique: Found[] = [];
  for (const entry of found) {
    if (!unique.some((kept) => sameControl(kept.control, entry.control))) {
      unique.push(entry);
    }
  }
  return unique
    .sort((a, b) => b.depth - a.depth || a.control.pos - b.control.pos)
    .map((entry) => entry.control);
}

function startOf(tr: Transaction, state: EditorState): number {
  const [first] = tr.steps;
  const reach = first ? reachesOf(first, state.doc)[0] : undefined;
  return reach?.from ?? state.selection.from;
}

/**
 * What the guards say about a transaction they turn down, or null when they let it through, which
 * also leaves out a transaction a plugin of the application's own filtered.
 */
export function editRefusal(
  tr: Transaction,
  state: EditorState
): EditRefusal | null {
  const reason = transactionRefusal(tr, state);
  if (reason === null) return null;
  return {
    reason,
    action: actionOf(tr),
    pos: startOf(tr, state),
    controls: reason === "lock" ? controlsOf(tr, state) : [],
  };
}
