/**
 * Evaluates OOXML content and deletion locks for inline controls and for the containers a control
 * stands around whole: a table cell, and a block-level control.
 *
 * `lockGuard` is what `./guards` registers all of this as, and it is the only way in: a caller
 * asking about a lock asks the one guard list, which asks the locks along with every other rule an
 * edit is judged by. What a lock says about a whole document or a whole selection - a question
 * about the document rather than about an edit - is exported beside it.
 */

import type {
  Fragment,
  Mark,
  Node as PMNode,
  ResolvedPos,
} from "prosemirror-model";
import { PluginKey, type Selection } from "prosemirror-state";
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
import type { EditIntent, StepGuard } from "./editGuard";
import { docxSchema } from "./index";
import { wrappersOf } from "./wrappers";

/**
 * The pass that lets a transaction through the guard, which is how a lock can be lifted at all.
 * A plugin key is used as the name so that it cannot collide with a consumer's own metadata.
 */
export const unlockAllowed = new PluginKey<boolean>("docxEditorUnlockAllowed");

/**
 * The pass an undo or a redo carries (`editor/commands/historyCommands`).
 *
 * Every step the history replays is the reverse of a step that passed the guard when it was made,
 * so replaying one leads back to a state that was allowed and nothing else. Without the pass the
 * reverse of a lock - a step across the very stretch that lock shut - would be refused, and the
 * refusal would take the whole history behind it down as well.
 */
export const historyReplay = new PluginKey<boolean>("docxEditorHistoryReplay");

/** What one control states about editing and deleting it, as the schema records it */
interface Locks {
  /** Whether its `w:lock` says the contents may not be edited */
  contents: boolean;
  /** Whether its `w:lock` says it may not be deleted, not even whole */
  deletion: boolean;
  /** Whether it is a `w:group`, which shuts its contents on terms of its own */
  group: boolean;
}

const OPEN: Locks = { contents: false, deletion: false, group: false };

/**
 * Whether this control shuts what stands inside it, where `inner` says whether another control
 * stands between that spot and this one: a `w:group` is superseded by one and a `w:lock` is not
 * (`spec/notes/contentControls.md`).
 */
function shutsContents(locks: Locks, inner: boolean): boolean {
  return locks.contents || (locks.group && !inner);
}

/** Whether this control shuts its contents with nothing inside it to supersede that */
function shutsAlone(locks: Locks): boolean {
  return shutsContents(locks, false);
}

interface StepRange {
  from: number;
  to: number;
}

/**
 * Every content control this inline node sits inside, locked or open, outermost first. Empty when
 * it sits in none.
 *
 * A control may hold a control (`schema/wrappers`), and each of them locks on its own terms.
 */
function sdtMarksOf(node: PMNode | null | undefined): readonly Mark[] {
  return node ? wrappersOf(node, docxSchema.marks.sdt) : [];
}

/** What the control this mark stands for states (`schema`) */
function markLocks(mark: Mark): Locks {
  return {
    contents: mark.attrs.contentsLocked === true,
    deletion: mark.attrs.deletionLocked === true,
    group: mark.attrs.group === true,
  };
}

/**
 * The control shutting the contents this inline node is part of. Null when it sits in no control,
 * or in none that shuts, a control locked against deletion alone included.
 *
 * This is about the node rather than about a spot inside it, so nothing here supersedes a group.
 */
export function lockedMarkOf(node: PMNode | null | undefined): Mark | null {
  return sdtMarksOf(node).find((mark) => shutsAlone(markLocks(mark))) ?? null;
}

/**
 * Every control a spot stands inside, outermost first.
 *
 * Only a spot with the very same control on both sides is inside it. At either edge the other side
 * belongs to a different control or to none, and since the mark is not inclusive what goes in
 * there falls outside the control.
 */
function controlsAcross($pos: ResolvedPos): readonly Mark[] {
  const after = sdtMarksOf($pos.nodeAfter);
  return sdtMarksOf($pos.nodeBefore).filter((mark) =>
    after.some((other) => other.eq(mark))
  );
}

/** Whether any of the controls a spot stands inside shuts it, the innermost one standing last */
function marksShut(marks: readonly Mark[]): boolean {
  return marks.some((mark, index) =>
    shutsContents(markLocks(mark), index < marks.length - 1)
  );
}

/** The attributes one kind of container writes what its control states in (`schema`) */
export interface LockAttrNames {
  contents: string;
  deletion: string;
  group: string;
}

/** A cell a control wraps carries them under names of its own, beside the cell's own attrs */
const CELL_ATTRS: LockAttrNames = {
  contents: "sdtContentsLocked",
  deletion: "sdtDeletionLocked",
  group: "sdtGroup",
};

/** A block-level control is the control, so it carries them under the wrapper's own names */
const BLOCK_ATTRS: LockAttrNames = {
  contents: "contentsLocked",
  deletion: "deletionLocked",
  group: "group",
};

/**
 * Which attributes the control standing at this node writes what it states in, and null where no
 * control stands there.
 *
 * These are the two containers a control stands around whole: a cell the file wrapped
 * (`docx/importTable`), which carries the control's opening XML beside its own attributes, and a
 * block-level control, which is the wrapper itself (`docx/importSdtBlock`). A cell no control
 * wrapped carries none of this, and counting it as a control would end the walk out of the tree at
 * the first cell and open every lock standing around the table.
 * A cell is found by its table role, so a schema built beside this one
 * (`table/__testing__/tables`) is read as well.
 */
export function lockAttrsOf(
  node: PMNode | null | undefined
): LockAttrNames | null {
  if (!node) return null;
  if (node.type.name === docxSchema.nodes.sdtBlock.name) return BLOCK_ATTRS;
  if (node.type.spec.tableRole !== "cell") return null;
  return typeof node.attrs.sdtPrefix === "string" ? CELL_ATTRS : null;
}

/** What the control this container stands for states. Nothing at all for anything else */
function containerLocks(node: PMNode | null | undefined): Locks {
  const names = lockAttrsOf(node);
  if (!node || !names) return OPEN;
  return {
    contents: node.attrs[names.contents] === true,
    deletion: node.attrs[names.deletion] === true,
    group: node.attrs[names.group] === true,
  };
}

/** Whether the control this container stands for shuts its contents (`schema`) */
export function isLockedContainer(node: PMNode | null | undefined): boolean {
  return shutsAlone(containerLocks(node));
}

/**
 * Whether this node carries a lock of either clause, an inline one wearing a control's mark and a
 * container standing for one alike.
 * This is the question about the document holding a lock at all rather than about editing a spot,
 * so a control locked against deletion alone counts (`editor/commands/lockCommands`). A `w:group`
 * is not a lock and does not: it shuts its contents, and there is nothing in it to lift.
 */
export function carriesLock(node: PMNode): boolean {
  const marks = node.isInline ? sdtMarksOf(node) : [];
  if (marks.length === 0) {
    const locks = containerLocks(node);
    return locks.contents || locks.deletion;
  }
  return marks.some((mark) => {
    const locks = markLocks(mark);
    return locks.contents || locks.deletion;
  });
}

/** A textblock a judgement or an edit runs through, and where its content begins */
export interface Textblock {
  node: PMNode;
  start: number;
}

/** One content control inside a textblock, as the one whole stretch it covers */
export interface ControlSpan extends StepRange {
  mark: Mark;
}

/**
 * Every control in this textblock, locked or open, each as the one whole stretch it covers.
 *
 * A control that wrapped several runs comes in as several inlines wearing the very same mark, and
 * what a lock answers for is the control rather than the run: a stretch covering one run of a
 * control whole still covers only a part of the control.
 *
 * A control standing inside another gives a span of its own, the outer one first, and the two
 * overlap: each locks on its own terms.
 */
export function controlSpans(block: Textblock): ControlSpan[] {
  const spans: ControlSpan[] = [];
  // The controls the node before this one stood in, so that a control reaching on is extended
  // rather than started again. Two controls never wear the same mark, so equality is the match
  let reaching: ControlSpan[] = [];
  block.node.forEach((child, offset) => {
    const from = block.start + offset;
    const to = from + child.nodeSize;
    const standing: ControlSpan[] = [];
    for (const mark of sdtMarksOf(child)) {
      const carried = reaching.find(
        (span) => span.to === from && span.mark.eq(mark)
      );
      if (carried) {
        carried.to = to;
        standing.push(carried);
        continue;
      }
      const started: ControlSpan = { from, to, mark };
      spans.push(started);
      standing.push(started);
    }
    reaching = standing;
  });
  return spans;
}

/**
 * What the container shutting this position holds, as the stretch a step has to reach to change
 * it. Null where no container around the position shuts it.
 *
 * `inner` says whether a control already stands between the position and the containers around it.
 * The walk runs outward and remembers every container it passes, since a container that does not
 * shut is still a control standing inside the next one out.
 */
function lockedContainerContent(
  $pos: ResolvedPos,
  inner: boolean
): StepRange | null {
  let passed = inner;
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth);
    if (lockAttrsOf(node) === null) continue;
    if (shutsContents(containerLocks(node), passed)) {
      const from = $pos.before(depth) + 1;
      return { from, to: from + node.content.size };
    }
    passed = true;
  }
  return null;
}

/** The same, for a spot whose own controls - the inline marks it stands inside - count as well */
function lockedContentAt(doc: PMNode, pos: number): StepRange | null {
  const $pos = doc.resolve(pos);
  return lockedContainerContent($pos, controlsAcross($pos).length > 0);
}

/**
 * Whether what stands at this position sits inside a container a control shuts.
 *
 * This is what the block intent is answered by, which is how a paragraph edit leaves its locked
 * paragraphs out (`editor/paragraphEdits`), and it is deliberately not `rangeTouchesLocked` over
 * the paragraph: a paragraph merely holding a locked control keeps its own alignment and indent.
 * An inline control the paragraph holds stands inside the paragraph rather than around it, so it
 * supersedes nothing here.
 */
function insideLockedContainer(doc: PMNode, pos: number): boolean {
  return lockedContainerContent(doc.resolve(pos), false) !== null;
}

/**
 * Whether this stretch reaches what a locked container holds.
 *
 * A stretch that only runs up to where the content begins or ends leaves that content alone: that
 * is the shape of a change to the container itself rather than to its contents, which is how a
 * column is given a new width. A stretch of no length at all, an insertion, is inside as soon as
 * it stands within the content.
 * A stretch covering the whole container reaches none of this, since it stands outside the content
 * on both sides: taking the container away whole is the deletion clause's question instead.
 */
function reachesLockedContainer(doc: PMNode, range: StepRange): boolean {
  const content =
    lockedContentAt(doc, range.from) ?? lockedContentAt(doc, range.to);
  if (!content) return false;
  if (range.from === range.to) return true;
  return range.from < content.to && range.to > content.from;
}

/** The stretch a step rewrites, and what it does to what stands there */
interface EditedRange extends StepRange {
  /**
   * Whether the step takes what stands in the stretch away, as against marking it where it
   * stands. Only a step that takes content away can be a control being deleted; a mark laid
   * across a whole control leaves the control standing, so it is an edit of its contents.
   */
  takesAway: boolean;
}

/** Whether the stretch covers this control from end to end, which is the control going whole */
function coversWhole(range: StepRange, span: StepRange): boolean {
  return range.from <= span.from && range.to >= span.to;
}

function overlaps(range: StepRange, span: StepRange): boolean {
  return range.from < span.to && range.to > span.from;
}

/**
 * Whether one control refuses this step, which is the two-question judgement.
 *
 * Covering the control whole and taking it away is the control being deleted as one, and the
 * deletion clause is the whole of the answer: the contents going with it is what a deletion is.
 * Anything else - a partial overlap, a mark laid over the control, or an insertion, which has no
 * length and so covers nothing whole - reaches into the contents, and the contents clause answers.
 */
function shuts(
  locks: Locks,
  range: EditedRange,
  span: StepRange,
  inner: boolean
): boolean {
  return range.takesAway && coversWhole(range, span)
    ? locks.deletion
    : shutsContents(locks, inner) && overlaps(range, span);
}

/** Whether a control inside this one holds the whole stretch, which is what supersedes a group */
function supersededSpan(
  spans: readonly ControlSpan[],
  span: ControlSpan,
  range: StepRange
): boolean {
  return spans.some(
    (inner) =>
      inner !== span && coversWhole(span, inner) && coversWhole(inner, range)
  );
}

/**
 * Whether any control this stretch meets refuses the step, each control judged by how much of it
 * the stretch covers.
 *
 * A container a control stands around is met here as the one whole thing it is, which is the shape
 * a row or column deletion writes, and the shape a block control is taken away in. A stretch that
 * merely runs into the container reaches its contents or nothing at all, and that is
 * `reachesLockedContainer`'s question: the two ends of a cell are what a change to the cell itself
 * covers, a new column width above all.
 * The outermost container the stretch covers whole is the one that answers, since the walk stops
 * at the first refusal; a stretch it lets through is then judged again against everything inside.
 */
function rangeShut(doc: PMNode, range: EditedRange): boolean {
  if (reachesLockedContainer(doc, range)) return true;
  let shut = false;
  doc.nodesBetween(range.from, range.to, (node, pos) => {
    if (shut) return false;
    if (lockAttrsOf(node) !== null) {
      const locks = containerLocks(node);
      if (coversWhole(range, { from: pos, to: pos + node.nodeSize })) {
        // The stretch stands outside the container, so nothing inside it supersedes a group
        if (range.takesAway ? locks.deletion : shutsAlone(locks)) shut = true;
      }
      // What the container holds is judged on its own, each control by its own terms
      return !shut;
    }
    if (!node.isTextblock) return true;
    const spans = controlSpans({ node, start: pos + 1 });
    // Never assigned, so a later textblock finding nothing cannot take an earlier refusal back
    shut ||= spans.some((span) =>
      shuts(
        markLocks(span.mark),
        range,
        span,
        supersededSpan(spans, span, range)
      )
    );
    return false;
  });
  return shut;
}

/**
 * Whether this stretch reaches contents a lock shuts: text inside a control that shuts its
 * contents, or what a locked container holds, the container itself included.
 *
 * This is the question about editing what stands there rather than about taking it away, so every
 * control the stretch meets answers with its contents clause. It is what the mark intent is
 * answered by, which is the stretch a character, link or lock edit is about to mark.
 */
function rangeTouchesLocked(doc: PMNode, from: number, to: number): boolean {
  return rangeShut(doc, { from, to, takesAway: false });
}

/**
 * Whether something inserted at this spot would land inside a locked control, whether it stands as
 * a mark or as a container.
 */
function insertionInsideLocked(doc: PMNode, pos: number): boolean {
  const $pos = doc.resolve(pos);
  const inside = controlsAcross($pos);
  return (
    marksShut(inside) ||
    lockedContainerContent($pos, inside.length > 0) !== null
  );
}

/** Whether a lock shuts editing what stands in this stretch, where it stands */
function markShut(doc: PMNode, from: number, to: number): boolean {
  return from === to
    ? insertionInsideLocked(doc, from)
    : rangeTouchesLocked(doc, from, to);
}

/** Whether a lock shuts putting something in place of what stands in this stretch */
function replaceShut(doc: PMNode, from: number, to: number): boolean {
  return from === to
    ? insertionInsideLocked(doc, from)
    : rangeShut(doc, { from, to, takesAway: true });
}

/**
 * Whether a lock shuts editing where this selection stands, which is what a menu that would lift a
 * lock is built on (`editor/commands/lockCommands`).
 *
 * This is a reading of the selection rather than of an edit - the lock guard answers an edit
 * (`./guards`) - and it names the locks alone, which is what the menu says. A caret is shut by
 * standing inside a control rather than against its edge, which is the same rule an insertion
 * follows.
 * A control locked against deletion alone shuts nothing here: its contents stand open, so editing
 * where it stands goes through, and only taking the control away is refused.
 */
export function selectionShut(selection: Selection, doc: PMNode): boolean {
  return selection.ranges.some((range) =>
    markShut(doc, range.$from.pos, range.$to.pos)
  );
}

/**
 * The stretches a step really rewrites, and what it does to each.
 *
 * A step working around a gap (`ReplaceAroundStep`) puts back what stands in the gap untouched,
 * so only its two ends count. That is what lets a paragraph holding a locked control still be
 * given a new alignment or indent, both of which rewrite the paragraph around its content.
 * Null for a step that has no such stretch.
 */
function editedRanges(step: Step): EditedRange[] | null {
  if (step instanceof ReplaceAroundStep) {
    return [
      { from: step.from, to: step.gapFrom, takesAway: true },
      { from: step.gapTo, to: step.to, takesAway: true },
    ];
  }
  if (step instanceof ReplaceStep) {
    return [{ from: step.from, to: step.to, takesAway: true }];
  }
  if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
    return [{ from: step.from, to: step.to, takesAway: false }];
  }
  return null;
}

function rangeAllowed(doc: PMNode, range: EditedRange): boolean {
  return range.from < range.to
    ? !rangeShut(doc, range)
    : !insertionInsideLocked(doc, range.from);
}

/** Whether this node would carry a lock into wherever the slice it stands in lands */
function plantedLock(node: PMNode): boolean {
  if (node.isInline) return lockedMarkOf(node) !== null;
  // A locked container with nothing inside it is the carrier of a change to the container's own
  // attributes (`setNodeMarkup`), which puts no content anywhere
  return isLockedContainer(node) && node.content.size > 0;
}

function lockedInside(content: Fragment): boolean {
  let found = false;
  content.descendants((node) => {
    if (found) return false;
    if (plantedLock(node)) found = true;
    return !found;
  });
  return found;
}

/**
 * Whether the step plants locked content somewhere.
 * A copy made inside the editor - dragging a stretch of text with Alt held - puts in the very
 * slice it was taken from and deletes nothing, so where it lands tells us nothing.
 */
function plantsLocked(step: Step): boolean {
  const carried =
    step instanceof ReplaceStep || step instanceof ReplaceAroundStep
      ? step.slice.content
      : null;
  return carried !== null && lockedInside(carried);
}

/**
 * Whether the step puts down one of the things a container's control states.
 *
 * What a cell records about itself is not its contents and stays open, which is what lets a column
 * holding a locked cell still be given a new width. What the control states is the one thing a
 * container records that a step of that very shape may not touch, so that lifting a lock stays the
 * business of the command that carries the pass for it, and a `w:group` cannot be put down at all.
 */
function clearsContainerLock(step: Step, doc: PMNode): boolean {
  if (step instanceof AttrStep) {
    if (step.value === true) return false;
    const at = doc.nodeAt(step.pos);
    const names = lockAttrsOf(at);
    if (!names) return false;
    const locks = containerLocks(at);
    if (step.attr === names.contents) return locks.contents;
    if (step.attr === names.deletion) return locks.deletion;
    if (step.attr === names.group) return locks.group;
    return false;
  }
  if (!(step instanceof ReplaceAroundStep)) return false;
  // The shape `setNodeMarkup` writes: the container standing here is replaced by one built afresh
  const was = containerLocks(doc.nodeAt(step.from));
  const now = containerLocks(step.slice.content.firstChild);
  return (
    (was.contents && !now.contents) ||
    (was.deletion && !now.deletion) ||
    (was.group && !now.group)
  );
}

function stepAllowed(step: Step, doc: PMNode): boolean {
  if (plantsLocked(step) || clearsContainerLock(step, doc)) return false;
  const ranges = editedRanges(step);
  if (ranges) return ranges.every((range) => rangeAllowed(doc, range));
  // A step that rewrites one node where it stands. Nothing in the package writes one today - a
  // new image size replaces the image node instead - so this stands for a consumer's own step
  if (
    step instanceof AttrStep ||
    step instanceof AddNodeMarkStep ||
    step instanceof RemoveNodeMarkStep
  ) {
    return (
      lockedContainerContent(doc.resolve(step.pos), false) === null &&
      lockedMarkOf(doc.nodeAt(step.pos)) === null
    );
  }
  // Anything else - a document attribute, a step a consumer brought - covers no content to protect
  return true;
}

/** Whether a lock shuts what this intent means to do where it stands */
function intentShut(doc: PMNode, intent: EditIntent): boolean {
  switch (intent.kind) {
    case "insert":
      return insertionInsideLocked(doc, intent.at);
    case "block":
      return insideLockedContainer(doc, intent.at);
    case "mark":
      return markShut(doc, intent.from, intent.to);
    case "replace":
      return replaceShut(doc, intent.from, intent.to);
    default: {
      const unmodelled: never = intent;
      return unmodelled;
    }
  }
}

/**
 * The locks the document carries, as `./guards` registers them.
 *
 * Both passes lift it: unlocking is the one edit that may reach into a lock, and every step the
 * history replays is the reverse of a step that passed the guard when it was made.
 */
export const lockGuard: StepGuard = {
  name: "lock",
  liftedBy: [unlockAllowed, historyReplay],
  step: (step, before) => stepAllowed(step, before),
  shuts: (intent, state) => intentShut(state.doc, intent),
};
