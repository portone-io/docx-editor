/**
 * How a display value gets onto a node: who says what it should be, and why the transaction that
 * writes it is no edit.
 *
 * What a node draws with is worked out from its source and the formatting around it
 * (`./attrRoles`), and has to be worked out again whenever either moves: the lines of a table's
 * cells once a row is added, the style values of a paragraph once its own properties change, and
 * every value at once when the formatting the document is resolved against is replaced. A deriver
 * says, for the node types it answers for, what the display attrs of one such node should be, and
 * `deriveDisplay` walks the blocks of a document, asks each owner in turn, and writes one step per
 * node whose attrs would change. Only the display attrs of what a deriver hands back are read, so
 * a deriver cannot write a source attr, and the walk is by construction the kind of transaction
 * `changesOnlyDisplayAttrs` describes.
 *
 * The interface stands here, below `docx` and `editor`, so that a module of either can write a
 * deriver against it. What a deriver resolves against is the caller's own `Context`, which this
 * module never reads; `editor/plugins/displayDerivation` registers the derivers and runs the walk
 * after every edit.
 *
 * A transaction carrying the `displayOnly` pass goes through the guard list whole (`./guards`),
 * provided every step of it is judged, off the role table alone, to change display attrs and
 * nothing else. The pass is a claim rather than a key: a step that rewrites a source attr, a lock
 * flag among them, or that puts content anywhere, fails the claim, and the transaction is then
 * judged as any other edit.
 */

import {
  type Attrs,
  Mark,
  type MarkType,
  type NodeType,
  type Node as PMNode,
  type Schema,
} from "prosemirror-model";
import { PluginKey } from "prosemirror-state";
import {
  AddMarkStep,
  AttrStep,
  ReplaceAroundStep,
  type Step,
  type Transform,
} from "prosemirror-transform";
import { attrRole } from "./attrRoles";
import type { docxSchema } from "./docxSchema";

/** The name of a node type of the document schema, which is what a deriver answers for by */
export type NodeName =
  typeof docxSchema extends Schema<infer Nodes, string> ? Nodes : never;

/** The display attrs one node should carry, by the position the node stands at in the document handed to the deriver */
export interface DisplayAttrs {
  readonly pos: number;
  readonly attrs: Attrs;
  /** Omitted for node attrs. Otherwise updates an existing mark on this inline node. */
  readonly mark?: { readonly type: MarkType; readonly to: number };
}

/**
 * What says, for the node types it answers for, what one node's display attrs should be.
 *
 * `derive` is asked once per such node, in document order, with the node as the derivers before it
 * left it. It hands back the display attrs of that node and of any node inside it, each by
 * position - a table's deriver answers for the cells - or nothing where nothing changes; a source
 * or session attr among what it hands back is not written. `previous` is the node this one maps
 * back to in the document before the change, or null when there is none, or when every value is
 * being worked out again from nothing.
 */
export interface DisplayDeriver<Context> {
  readonly name: string;
  /**
   * The block node types this deriver answers for. The walk runs over the blocks and never into
   * the inline content of a textblock, since it runs after every keystroke, so an inline type
   * cannot be answered for.
   */
  readonly nodeTypes: readonly NodeName[];
  derive(
    node: PMNode,
    pos: number,
    doc: PMNode,
    context: Context,
    previous: PMNode | null
  ): readonly DisplayAttrs[];
}

/** The derivers answering for each node type, in registration order */
function ownersByType<Context>(
  schema: Schema,
  derivers: readonly DisplayDeriver<Context>[]
): Map<string, DisplayDeriver<Context>[]> {
  const owners = new Map<string, DisplayDeriver<Context>[]>();
  for (const deriver of derivers) {
    for (const name of deriver.nodeTypes) {
      const type = schema.nodes[name];
      if (type === undefined || type.isInline) {
        throw new Error(
          `${deriver.name} answers for ${name}, which the walk over the blocks never reaches`
        );
      }
      const owning = owners.get(name) ?? [];
      owning.push(deriver);
      owners.set(name, owning);
    }
  }
  return owners;
}

/** The display attrs among what a deriver handed back, laid over the attrs the node standing there carries */
function displayLaidOver(standing: PMNode | Mark, derived: Attrs): Attrs {
  const attrs: Record<string, unknown> = { ...standing.attrs };
  for (const [name, value] of Object.entries(derived)) {
    if (attrRole(standing.type, name) === "display") attrs[name] = value;
  }
  return attrs;
}

function nodeStanding(doc: PMNode, pos: number): PMNode {
  const node = doc.nodeAt(pos);
  if (node === null) throw new Error(`no node stands at ${pos}`);
  return node;
}

/**
 * Writes the display attrs the derivers hand back, one `setNodeMarkup` per node whose attrs would
 * change, onto the transform. It holds no state of its own: what a node maps back to comes from
 * the caller, who knows what changed.
 *
 * Every step keeps the document the same size, so the positions gathered before the first write
 * still hold, and a deriver asked after another reads the node as that one left it.
 */
export function deriveDisplay<Context>(
  transform: Transform,
  context: Context,
  derivers: readonly DisplayDeriver<Context>[],
  previousOf: (node: PMNode, pos: number) => PMNode | null
): void {
  const owners = ownersByType(transform.doc.type.schema, derivers);
  const spots: { pos: number; owners: readonly DisplayDeriver<Context>[] }[] =
    [];
  transform.doc.descendants((node, pos) => {
    const owning = owners.get(node.type.name);
    if (owning !== undefined) spots.push({ pos, owners: owning });
    return !node.isTextblock;
  });
  for (const spot of spots) {
    for (const deriver of spot.owners) {
      const node = nodeStanding(transform.doc, spot.pos);
      const derived = deriver.derive(
        node,
        spot.pos,
        transform.doc,
        context,
        previousOf(node, spot.pos)
      );
      for (const { pos, attrs, mark: target } of derived) {
        const standing = nodeStanding(transform.doc, pos);
        if (target !== undefined) {
          const mark = target.type.isInSet(standing.marks);
          if (!standing.isInline || mark === undefined) continue;
          const end =
            pos + standing.nodeSize - transform.doc.resolve(pos).textOffset;
          if (target.to <= pos || target.to > end) {
            throw new Error(
              "a display mark update must stay within its inline node"
            );
          }
          const next = target.type.create(displayLaidOver(mark, attrs));
          if (!mark.eq(next)) {
            // AddMarkStep replaces this type atomically. Transform.addMark would first remove
            // the source-bearing mark, which is correctly refused by the display-only gate.
            transform.step(new AddMarkStep(pos, target.to, next));
          }
          continue;
        }
        const next = displayLaidOver(standing, attrs);
        if (!standing.hasMarkup(standing.type, next, standing.marks)) {
          transform.setNodeMarkup(pos, null, next);
        }
      }
    }
  }
}

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
function sameOutsideDisplay(
  type: NodeType | MarkType,
  was: Attrs,
  now: Attrs
): boolean {
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
 * Node attrs and existing mark attrs may change only their display fields. Adding a mark where
 * that type was absent, changing source attrs, and all other step shapes remain edits.
 */
export function changesOnlyDisplayAttrs(step: Step, doc: PMNode): boolean {
  if (step instanceof AddMarkStep) {
    let touched = false;
    let allowed = true;
    doc.nodesBetween(step.from, step.to, (node, _pos, parent) => {
      if (!node.isInline) return true;
      if (!parent?.type.allowsMarkType(step.mark.type)) return false;
      touched = true;
      const previous = step.mark.type.isInSet(node.marks);
      if (
        !previous ||
        !sameOutsideDisplay(step.mark.type, previous.attrs, step.mark.attrs)
      )
        allowed = false;
      return false;
    });
    return touched && allowed;
  }
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
