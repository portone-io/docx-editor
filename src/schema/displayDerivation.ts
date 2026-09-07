/**
 * A transaction that changes nothing but display values.
 *
 * What a node draws with is worked out from its source and the formatting around it
 * (`./attrRoles`), and a step that only writes such a value again edits nothing: no lock shuts a
 * display value and no protection holds one back. So a transaction carrying the `displayOnly` pass
 * goes through the guard list whole (`./guards`), provided every step of it is judged, off the
 * role table alone, to change display attrs and nothing else.
 *
 * The pass is a claim rather than a key. A step that rewrites a source attr, a lock flag among
 * them, or that puts content anywhere, fails the claim, and the transaction is then judged as any
 * other edit; the claim only spares an honest re-derivation the guards it could not pass.
 */

import {
  type Attrs,
  Mark,
  type NodeType,
  type Node as PMNode,
} from "prosemirror-model";
import { PluginKey } from "prosemirror-state";
import { AttrStep, ReplaceAroundStep, type Step } from "prosemirror-transform";
import { attrRole } from "./attrRoles";

/**
 * The pass a re-derivation carries. A plugin key is used as the name so that it cannot collide
 * with a consumer's own metadata.
 */
export const displayOnly = new PluginKey<boolean>("docxEditorDisplayOnly");

/**
 * Whether the two attr sets agree on every attr but the display ones of this type.
 *
 * A source or session attr has to be the very same value: a re-derivation spreads the node's own
 * attrs and changes none of them, and a value rebuilt to read the same is reported as a change
 * rather than hidden, which is what the rest of the guard list is for.
 */
function sameOutsideDisplay(type: NodeType, was: Attrs, now: Attrs): boolean {
  const names = new Set([...Object.keys(was), ...Object.keys(now)]);
  for (const name of names) {
    if (attrRole(type, name) === "display") continue;
    if (was[name] !== now[name]) return false;
  }
  return true;
}

/**
 * The node standing where the step rewrites one, and the node the step puts in its place, when
 * the step has the shape `setNodeMarkup` writes: the one node whole, its content kept as the gap.
 * Null for a step of any other shape.
 */
function rewrittenNode(
  step: ReplaceAroundStep,
  doc: PMNode
): { was: PMNode; now: PMNode } | null {
  const was = doc.nodeAt(step.from);
  const now = step.slice.content.firstChild;
  if (
    was === null ||
    now === null ||
    step.slice.content.childCount !== 1 ||
    step.slice.openStart !== 0 ||
    step.slice.openEnd !== 0 ||
    now.content.size !== 0 ||
    step.insert !== 1 ||
    step.gapFrom !== step.from + 1 ||
    step.to !== step.from + was.nodeSize ||
    step.gapTo !== step.to - 1
  ) {
    return null;
  }
  return { was, now };
}

/**
 * Whether the step changes display attrs and nothing else, judged off the role table alone.
 *
 * Two shapes qualify: an attr step naming a display attr, and a node rewritten where it stands
 * (`setNodeMarkup`) as the same type wearing the same marks, differing in display attrs only. A
 * step of any other shape puts content somewhere or takes it away, and is an edit.
 */
export function changesOnlyDisplayAttrs(step: Step, doc: PMNode): boolean {
  if (step instanceof AttrStep) {
    const node = doc.nodeAt(step.pos);
    return node !== null && attrRole(node.type, step.attr) === "display";
  }
  if (!(step instanceof ReplaceAroundStep)) return false;
  const rewritten = rewrittenNode(step, doc);
  if (rewritten === null) return false;
  const { was, now } = rewritten;
  return (
    was.type === now.type &&
    Mark.sameSet(was.marks, now.marks) &&
    sameOutsideDisplay(was.type, was.attrs, now.attrs)
  );
}
