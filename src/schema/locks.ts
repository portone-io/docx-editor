/**
 * Evaluates OOXML content and deletion locks for inline controls and whole table cells. Commands
 * query these predicates before editing; `editor/plugins/lockedContent` enforces them at runtime.
 */

import type {
  Fragment,
  Mark,
  Node as PMNode,
  ResolvedPos,
} from "prosemirror-model";
import { PluginKey, type Selection, type Transaction } from "prosemirror-state";
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
import { docxSchema } from "./index";

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

/** The two clauses a control's lock settles, as the schema records them */
interface Locks {
  /** Whether the contents may not be edited */
  contents: boolean;
  /** Whether the control may not be deleted, not even whole */
  deletion: boolean;
}

const OPEN: Locks = { contents: false, deletion: false };

interface StepRange {
  from: number;
  to: number;
}

/** The content control this inline node sits inside, locked or open. Null when it sits in none */
function sdtMarkOf(node: PMNode | null | undefined): Mark | null {
  return node?.marks.find((mark) => mark.type === docxSchema.marks.sdt) ?? null;
}

/** What the control this mark stands for shuts (`schema`) */
function markLocks(mark: Mark): Locks {
  return {
    contents: mark.attrs.contentsLocked === true,
    deletion: mark.attrs.deletionLocked === true,
  };
}

/**
 * The control shutting the contents this inline node is part of. Null when it sits in no control,
 * or in one whose contents stand open, a control locked against deletion alone included.
 */
export function lockedMarkOf(node: PMNode | null | undefined): Mark | null {
  const mark = sdtMarkOf(node);
  return mark !== null && markLocks(mark).contents ? mark : null;
}

/** The two cell attributes the clauses of a wrapped cell's lock are written in (`schema`) */
const CELL_CONTENTS_ATTR = "sdtContentsLocked";
const CELL_DELETION_ATTR = "sdtDeletionLocked";

function isCell(node: PMNode | null | undefined): boolean {
  return node?.type.spec.tableRole === "cell";
}

/** What the control around this cell shuts. Nothing at all for anything that is not a cell */
function cellLocks(node: PMNode | null | undefined): Locks {
  if (!node || !isCell(node)) return OPEN;
  return {
    contents: node.attrs[CELL_CONTENTS_ATTR] === true,
    deletion: node.attrs[CELL_DELETION_ATTR] === true,
  };
}

/** Whether the control around this cell shuts its contents (`sdtContentsLocked` in `schema`) */
export function isLockedCell(node: PMNode | null | undefined): boolean {
  return cellLocks(node).contents;
}

/**
 * Whether this node carries a lock of either clause, an inline one wearing a control's mark and a
 * wrapped cell alike.
 * This is the question about the document holding a lock at all rather than about editing a spot,
 * so a control locked against deletion alone counts (`editor/commands/lockCommands`).
 */
export function carriesLock(node: PMNode): boolean {
  const mark = node.isInline ? sdtMarkOf(node) : null;
  const locks = mark === null ? cellLocks(node) : markLocks(mark);
  return locks.contents || locks.deletion;
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
 */
export function controlSpans(block: Textblock): ControlSpan[] {
  const spans: ControlSpan[] = [];
  block.node.forEach((child, offset) => {
    const mark = sdtMarkOf(child);
    if (!mark) return;
    const from = block.start + offset;
    const to = from + child.nodeSize;
    const last = spans.at(-1);
    if (last && last.to === from && last.mark.eq(mark)) last.to = to;
    else spans.push({ from, to, mark });
  });
  return spans;
}

/** What a locked cell holds, as the stretch a step has to reach to change it. Null outside every locked cell */
function lockedCellContent($pos: ResolvedPos): StepRange | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth);
    if (!isLockedCell(node)) continue;
    const from = $pos.before(depth) + 1;
    return { from, to: from + node.content.size };
  }
  return null;
}

/**
 * Whether what stands at this position sits inside a cell a control shuts.
 *
 * This is the question a paragraph edit leaves its locked paragraphs out by
 * (`editor/paragraphEdits`), and it is deliberately not `rangeTouchesLocked` over the paragraph:
 * a paragraph merely holding a locked control keeps its own alignment and indent.
 */
export function insideLockedCell(doc: PMNode, pos: number): boolean {
  return lockedCellContent(doc.resolve(pos)) !== null;
}

/**
 * Whether this stretch reaches what a locked cell holds.
 *
 * A stretch that only runs up to where the content begins or ends leaves that content alone: that
 * is the shape of a change to the cell itself rather than to its contents, which is how a column
 * is given a new width. A stretch of no length at all, an insertion, is inside as soon as it
 * stands within the content.
 * A stretch covering the whole cell reaches none of this, since it stands outside the content on
 * both sides: taking the cell away whole is the deletion clause's question instead.
 */
function reachesLockedCell(doc: PMNode, range: StepRange): boolean {
  const content =
    lockedCellContent(doc.resolve(range.from)) ??
    lockedCellContent(doc.resolve(range.to));
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
 * Whether the lock on one control refuses this step, which is the two-question judgement.
 *
 * Covering the control whole and taking it away is the control being deleted as one, and the
 * deletion clause is the whole of the answer: the contents going with it is what a deletion is.
 * Anything else - a partial overlap, a mark laid over the control, or an insertion, which has no
 * length and so covers nothing whole - reaches into the contents, and the contents clause answers.
 */
function shuts(locks: Locks, range: EditedRange, span: StepRange): boolean {
  return range.takesAway && coversWhole(range, span)
    ? locks.deletion
    : locks.contents && overlaps(range, span);
}

/**
 * Whether any control this stretch meets refuses the step, each control judged by how much of it
 * the stretch covers.
 *
 * A wrapped cell is met here as the one whole thing it is, which is the shape a row or column
 * deletion writes and the only shape either clause answers for. A stretch that merely runs into
 * the cell reaches its contents or nothing at all, and that is `reachesLockedCell`'s question: the
 * two ends of the cell are what a change to the cell itself covers, a new column width above all.
 */
function rangeShut(doc: PMNode, range: EditedRange): boolean {
  if (reachesLockedCell(doc, range)) return true;
  let shut = false;
  doc.nodesBetween(range.from, range.to, (node, pos) => {
    if (shut) return false;
    if (isCell(node)) {
      const locks = cellLocks(node);
      if (coversWhole(range, { from: pos, to: pos + node.nodeSize })) {
        if (range.takesAway ? locks.deletion : locks.contents) shut = true;
      }
      // What the cell holds is judged on its own, each control by its own lock
      return !shut;
    }
    if (!node.isTextblock) return true;
    // Never assigned, so a later textblock finding nothing cannot take an earlier refusal back
    shut ||= controlSpans({ node, start: pos + 1 }).some((span) =>
      shuts(markLocks(span.mark), range, span)
    );
    return false;
  });
  return shut;
}

/**
 * Whether this stretch reaches contents a lock shuts: text inside a control that shuts its
 * contents, or what a locked cell holds, the cell itself included.
 *
 * This is the question about editing what stands there rather than about taking it away, so every
 * control the stretch meets answers with its contents clause. The lock commands and the formatting
 * commands ask it of the stretch they are about to mark, so it is exported.
 */
export function rangeTouchesLocked(
  doc: PMNode,
  from: number,
  to: number
): boolean {
  return rangeShut(doc, { from, to, takesAway: false });
}

/**
 * Whether something inserted at this spot would land inside a locked control or a locked cell.
 *
 * Only a spot with the very same control on both sides is inside one. At either edge the other
 * side belongs to a different control or to none, and since the mark is not inclusive what goes
 * in there falls outside the control.
 */
export function insertionInsideLocked(doc: PMNode, pos: number): boolean {
  const $pos = doc.resolve(pos);
  if (lockedCellContent($pos)) return true;
  const before = lockedMarkOf($pos.nodeBefore);
  const after = lockedMarkOf($pos.nodeAfter);
  return before !== null && after !== null && before.eq(after);
}

/**
 * Whether the guard shuts editing where this selection stands.
 *
 * This is the question a command that rewrites whatever is selected has to ask before it reports
 * that it applies (`editor/insertImage`), and the one a menu offering to lift a lock is built on
 * (`editor/commands/lockCommands`). A caret is shut by standing inside a control rather than
 * against its edge, which is the same rule an insertion follows.
 * A control locked against deletion alone shuts nothing here: its contents stand open, so editing
 * where it stands goes through, and only taking the control away is refused.
 */
export function selectionShut(selection: Selection, doc: PMNode): boolean {
  return selection.ranges.some((range) =>
    range.$from.pos === range.$to.pos
      ? insertionInsideLocked(doc, range.$from.pos)
      : rangeTouchesLocked(doc, range.$from.pos, range.$to.pos)
  );
}

/**
 * Whether the guard would refuse putting something in place of what this selection covers, which
 * is what an image insertion does (`editor/insertImage`).
 *
 * A different question from `selectionShut`: what the selection covers whole goes away rather than
 * being edited, so the deletion clause answers for it. A control that may be taken away whole may
 * therefore be replaced, and one locked against deletion alone may not, even though editing inside
 * it would have gone through.
 */
export function replacementShut(selection: Selection, doc: PMNode): boolean {
  return selection.ranges.some((range) => {
    const from = range.$from.pos;
    const to = range.$to.pos;
    return from === to
      ? insertionInsideLocked(doc, from)
      : rangeShut(doc, { from, to, takesAway: true });
  });
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
  // A locked cell with nothing inside it is the carrier of a change to the cell's own attributes
  // (`setNodeMarkup`), which puts no content anywhere
  return isLockedCell(node) && node.content.size > 0;
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
 * Whether the step puts one of a cell's own locks down.
 *
 * What a cell records about itself is not its contents and stays open, which is what lets a column
 * holding a locked cell still be given a new width. The two locks are the one thing it records
 * that a step of that very shape may not touch, so that lifting a lock stays the business of the
 * command that carries the pass for it.
 */
function clearsCellLock(step: Step, doc: PMNode): boolean {
  if (step instanceof AttrStep) {
    if (step.value === true) return false;
    const locks = cellLocks(doc.nodeAt(step.pos));
    if (step.attr === CELL_CONTENTS_ATTR) return locks.contents;
    if (step.attr === CELL_DELETION_ATTR) return locks.deletion;
    return false;
  }
  if (!(step instanceof ReplaceAroundStep)) return false;
  // The shape `setNodeMarkup` writes: the cell standing here is replaced by one built afresh
  const was = cellLocks(doc.nodeAt(step.from));
  const now = cellLocks(step.slice.content.firstChild);
  return (was.contents && !now.contents) || (was.deletion && !now.deletion);
}

function stepAllowed(step: Step, doc: PMNode): boolean {
  if (plantsLocked(step) || clearsCellLock(step, doc)) return false;
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
      lockedCellContent(doc.resolve(step.pos)) === null &&
      lockedMarkOf(doc.nodeAt(step.pos)) === null
    );
  }
  // Anything else - a document attribute, a step a consumer brought - covers no content to protect
  return true;
}

/** Whether the transaction carries one of the two passes through the guard */
function carriesPass(tr: Transaction): boolean {
  return (
    tr.getMeta(unlockAllowed) === true || tr.getMeta(historyReplay) === true
  );
}

/**
 * Whether the guard would let this transaction through, decided and nothing else.
 *
 * The refusal the guard itself answers with carries a side effect - the composition it ends
 * (`editor/plugins/lockedContent`) - which a query about a button's state may not set off, so the
 * decision stands apart from it and every caller building an edit asks this rather than handing
 * the transaction to a state.
 * `doc` only stands in for a step the transaction kept no document for.
 */
export function transactionAllowed(tr: Transaction, doc: PMNode): boolean {
  if (!tr.docChanged) return true;
  if (carriesPass(tr)) return true;
  // Each step counts positions in the document it was built against, which `docs` holds
  return tr.steps.every((step, index) =>
    stepAllowed(step, tr.docs[index] ?? doc)
  );
}
