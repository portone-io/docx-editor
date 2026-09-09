/**
 * Keeps a comment when the text it was written for goes away.
 *
 * Deleting commented text is an edit of the body, not a decision about the thread: a reader who
 * rewrites a sentence has not answered what was said about it. So a body edit that takes a
 * comment's nodes away has them put back at the spots the edit left, and the thread carries on -
 * anchored where one of its two range markers survived, detached where both went with the text
 * (`DocumentComment.anchored`). What the comment and its replies say needs nothing put back: a body
 * is a story on the document node (`docx/story`), and a deletion in the body does not reach it.
 *
 * That story is also what tells this from taking a comment down on purpose, since `removeComment`
 * is the one edit that drops it. A comment is therefore put back only while its story stands, which
 * needs neither an instruction passed from the command nor a reading of the history: undo puts the
 * story back along with everything else, and redo of a removal drops it again. The cost is that a
 * reference the comments part held no body for is never put back, and a comment saying nothing is
 * nothing to keep.
 *
 * Two limits are worth naming. A restoration is an ordinary transaction and the guards judge it
 * (`editor/plugins/lockedContent`), so a spot no insertion is allowed at is no home: the nodes go
 * to the nearest spot the guards leave open, and a document leaving none anywhere - every textblock
 * inside a locked control - loses the comment. And a range marker with no home takes its surviving
 * counterpart down with it, so that no export carries half a range.
 *
 * An open IME composition needs no deferral of the kind `displayDerivation` makes. That plugin
 * rewrites the node being composed in, which is what takes a composition down; this one puts an
 * atom in beside the composed text and leaves the text itself alone, and Chrome carries the
 * composition on through it (`e2e/hangulComposition.spec.ts`).
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  Plugin,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { storyKey, storyOf } from "../../docx/story";
import { docxSchema } from "../../schema";
import { editShut, transactionTouchesComments } from "../../schema/guards";
import { stringAttr } from "../commands/comments/model";

/** The three nodes a comment stands in the story as, named as this module holds them */
type NodeField = "start" | "end" | "reference";

const NODE_FIELD: Readonly<Record<string, NodeField | undefined>> = {
  commentStart: "start",
  commentEnd: "end",
  commentReference: "reference",
};

/** Put back in this order, which is the order they stand in where they land on one spot */
const NODE_FIELDS: readonly NodeField[] = ["start", "end", "reference"];

/**
 * Which way the edit's mapping is asked to lean for each node: the opening marker follows the text
 * after it, the closing marker and the reference the text before them. A stretch trimmed at either
 * end therefore keeps its markers around what is left of it rather than around what the edit put
 * in its place.
 */
const MAPPING_SIDE: Readonly<Record<NodeField, 1 | -1>> = {
  start: 1,
  end: -1,
  reference: -1,
};

interface PlacedNode {
  node: PMNode;
  pos: number;
}

type CommentNodes = { [Field in NodeField]: PlacedNode | null };

const NO_NODES: CommentNodes = { start: null, end: null, reference: null };

/** The first of each of a comment's three nodes, by comment id */
function commentNodesIn(doc: PMNode): Map<string, CommentNodes> {
  const found = new Map<string, CommentNodes>();
  doc.descendants((node, pos) => {
    const field = NODE_FIELD[node.type.name];
    if (field === undefined) return true;
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    const nodes = found.get(id) ?? { ...NO_NODES };
    found.set(id, nodes);
    if (nodes[field] === null) nodes[field] = { node, pos };
    return true;
  });
  return found;
}

interface InlineRange {
  from: number;
  to: number;
}

/** Every stretch of the document a comment's nodes may stand in, in document order */
function inlineRanges(doc: PMNode): InlineRange[] {
  const type = docxSchema.nodes.commentReference;
  const ranges: InlineRange[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    if (node.type.contentMatch.matchType(type) !== null) {
      ranges.push({ from: pos + 1, to: pos + 1 + node.content.size });
    }
    return false;
  });
  return ranges;
}

function within(ranges: readonly InlineRange[], pos: number): boolean {
  return ranges.some((range) => pos >= range.from && pos <= range.to);
}

/**
 * The spots a node the edit deleted could go back to, best first: where the edit left it, then
 * where the deletion left the caret, then outwards from there, backward before forward.
 *
 * Backward first because what was deleted stood after the text that led up to it, so the reader
 * reads the thread next to what is left of its neighbourhood.
 */
function* homes(
  ranges: readonly InlineRange[],
  wanted: number,
  caret: number
): Generator<number> {
  if (within(ranges, wanted)) yield wanted;
  if (caret !== wanted && within(ranges, caret)) yield caret;
  for (let index = ranges.length - 1; index >= 0; index -= 1) {
    const range = ranges[index];
    if (range !== undefined && range.to < wanted) yield range.to;
  }
  for (const range of ranges) {
    if (range.from > wanted) yield range.from;
  }
}

/**
 * The first of those spots the guards leave open, and null where they leave none.
 *
 * The same question `openStretches` asks, and it has to be asked here: a transaction appended by a
 * plugin is judged by `filterTransaction` like any other and dropped without a word when it is
 * refused, so a home inside a locked control would lose the comment silently.
 */
function allowedHome(
  state: EditorState,
  ranges: readonly InlineRange[],
  wanted: number,
  caret: number
): number | null {
  for (const at of homes(ranges, wanted, caret)) {
    if (!editShut(state, { kind: "insert", at })) return at;
  }
  return null;
}

/** The position this one had before the edit, counted in the document the edit left */
function mappedThrough(
  transactions: readonly Transaction[],
  pos: number,
  side: 1 | -1
): number {
  return transactions.reduce((at, tr) => tr.mapping.map(at, side), pos);
}

/**
 * The node as it goes back in: its run mark kept, the wrappers it stood inside dropped.
 *
 * The run mark carries the character style Word draws a comment mark with, and a run mark is valid
 * wherever inline content is. A content control or a hyperlink is a stretch of the document rather
 * than something a node carries, and the stretch this one stood in may be part of what went.
 */
function restorable(node: PMNode): PMNode {
  return node.type.create(
    node.attrs,
    null,
    node.marks.filter((mark) => mark.type === docxSchema.marks.run)
  );
}

interface Insertion {
  node: PMNode;
  at: number;
  /** Where it stood before the edit, which orders two nodes landing on one spot */
  was: number;
}

interface Plan {
  insertions: Insertion[];
  /** A surviving marker whose counterpart has no home, which may not be written out alone */
  removals: PlacedNode[];
}

function planRestoration(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState
): Plan {
  const before = commentNodesIn(oldState.doc);
  const after = commentNodesIn(newState.doc);
  const ranges = inlineRanges(newState.doc);
  const caret = newState.selection.from;
  const plan: Plan = { insertions: [], removals: [] };
  for (const [id, was] of before) {
    if (storyOf(newState.doc, storyKey("comment", id)) === null) continue;
    const now = after.get(id) ?? NO_NODES;
    for (const field of NODE_FIELDS) {
      const lost = was[field];
      // Nothing to put back where none stood, or where a copy of it still does: a document can
      // hold one comment's nodes twice, since a copy carries the id it was written with
      if (lost === null || now[field] !== null) continue;
      const counterpart =
        field === "start" ? now.end : field === "end" ? now.start : null;
      // A comment that lost both markers has no stretch left to mark and carries on detached
      if (field !== "reference" && counterpart === null) continue;
      const at = allowedHome(
        newState,
        ranges,
        mappedThrough(transactions, lost.pos, MAPPING_SIDE[field]),
        caret
      );
      if (at !== null) {
        plan.insertions.push({
          node: restorable(lost.node),
          at,
          was: lost.pos,
        });
      } else if (counterpart !== null && !markerShut(newState, counterpart)) {
        plan.removals.push(counterpart);
      }
    }
  }
  return plan;
}

function markerShut(state: EditorState, marker: PlacedNode): boolean {
  return editShut(state, {
    kind: "replace",
    from: marker.pos,
    to: marker.pos + marker.node.nodeSize,
  });
}

/**
 * Where the caret goes when the edit turns out to have deleted comment nodes and nothing else.
 *
 * A range marker is invisible and takes no selection, so Backspace against one deletes the marker
 * and the restoration puts it straight back. Left where the mapping puts it, the caret would face
 * the same marker again and the key would be dead; carried on in the direction the deletion went,
 * past every comment node standing in the way, one press steps over a whole marker pair.
 */
function pastCommentNodes(doc: PMNode, pos: number, direction: 1 | -1): number {
  let at = pos;
  for (;;) {
    const resolved = doc.resolve(at);
    const next = direction < 0 ? resolved.nodeBefore : resolved.nodeAfter;
    if (next === null || NODE_FIELD[next.type.name] === undefined) return at;
    at += direction * next.nodeSize;
  }
}

function withCaretPastMarkers(
  tr: Transaction,
  oldState: EditorState,
  newState: EditorState
): Transaction {
  const { selection } = newState;
  if (!(selection instanceof TextSelection) || !selection.empty) return tr;
  if (!tr.doc.eq(oldState.doc)) return tr;
  const direction = selection.from < oldState.selection.from ? -1 : 1;
  return tr.setSelection(
    TextSelection.create(
      tr.doc,
      pastCommentNodes(tr.doc, tr.selection.from, direction)
    )
  );
}

function restoration(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState
): Transaction | null {
  const plan = planRestoration(transactions, oldState, newState);
  if (plan.insertions.length === 0 && plan.removals.length === 0) return null;
  const tr = newState.tr;
  for (const marker of [...plan.removals].sort(
    (left, right) => right.pos - left.pos
  )) {
    tr.delete(marker.pos, marker.pos + marker.node.nodeSize);
  }
  // Ascending, so that nodes swept away by one edit go back in the order they stood in. Each spot
  // is counted in the document before this transaction, so it is mapped through what the deletions
  // and the insertions ahead of it have already moved
  const ordered = [...plan.insertions].sort(
    (left, right) => left.at - right.at || left.was - right.was
  );
  for (const { node, at } of ordered) {
    tr.insert(tr.mapping.map(at), node);
  }
  if (!tr.docChanged) return null;
  return withCaretPastMarkers(tr, oldState, newState);
}

export function commentRestoration(): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      // An edit that leaves every comment node where it stood cannot have taken one away, and this
      // is the same reading `protectionAllowsTransaction` settles the common keystroke with
      if (!transactions.some(transactionTouchesComments)) return null;
      return restoration(transactions, oldState, newState);
    },
  });
}
