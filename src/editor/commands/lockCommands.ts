/**
 * Adds or removes OOXML content locks. Multi-paragraph selections become separate controls, and
 * unlocking preserves the control metadata while removing only its lock.
 */

import type { Attrs, Mark, Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { namesNothing, newControlId } from "../../docx/sdt";
import { withContentLock } from "../../docx/sdtProps";
import { docxSchema } from "../../schema";
import {
  type ControlSpan,
  carriesLock,
  controlSpans,
  isLockedCell,
  rangeTouchesLocked,
  selectionShut,
  type Textblock,
  unlockAllowed,
} from "../../schema/locks";
import { editsShut } from "../../schema/protectionState";

/**
 * The mark a newly locked stretch wears, which is also the XML it goes back out as.
 * `sdtContentLocked` shuts both clauses of the lock, so the mark records both.
 */
function lockedControlMark(id: number): Mark {
  return docxSchema.marks.sdt.create({
    sdtPrefix:
      "<w:sdt><w:sdtPr>" +
      `<w:id w:val="${id}"/><w:lock w:val="sdtContentLocked"/>` +
      "</w:sdtPr>",
    contentsLocked: true,
    deletionLocked: true,
  });
}

/** Whether this control shuts its contents, which is the lock these commands offer to lift */
function isLocked(mark: Mark): boolean {
  return mark.attrs.contentsLocked === true;
}

/** The opening XML a control goes back out as. Null for one that has lost it */
function prefixOf(attrs: Attrs): string | null {
  const prefix: unknown = attrs.sdtPrefix;
  return typeof prefix === "string" ? prefix : null;
}

/**
 * The same control with its lock shut or lifted, which is how a lock moves without the control
 * moving with it.
 * Both clauses move together, since what is written is either `sdtContentLocked`, which shuts
 * both, or no `w:lock` at all.
 * null for an opening we could not rewrite, which leaves the caller to decide what to do instead.
 */
function withLock(mark: Mark, locked: boolean): Mark | null {
  const prefix = prefixOf(mark.attrs);
  const next = prefix === null ? null : withContentLock(prefix, locked);
  if (next === null) return null;
  return mark.type.create({
    ...mark.attrs,
    sdtPrefix: next,
    contentsLocked: locked,
    deletionLocked: locked,
  });
}

interface Stretch {
  from: number;
  to: number;
}

/** Whether these two stretches share any length of text */
function overlaps(a: Stretch, b: Stretch): boolean {
  return a.from < b.to && a.to > b.from;
}

/** Whether this stretch reaches the other, an edge counting as reaching it */
function reaches(a: Stretch, b: Stretch): boolean {
  return a.from <= b.to && a.to >= b.from;
}

/** The part of this paragraph the selection covers. Null where it covers none of it */
function stretchIn(block: Textblock, reach: Stretch): Stretch | null {
  const from = Math.max(block.start, reach.from);
  const to = Math.min(block.start + block.node.content.size, reach.to);
  return from < to ? { from, to } : null;
}

/** The parts of this stretch that no control stands in */
function gapsIn(stretch: Stretch, spans: readonly ControlSpan[]): Stretch[] {
  const gaps: Stretch[] = [];
  let at = stretch.from;
  for (const span of spans) {
    const to = Math.min(span.from, stretch.to);
    if (at < to) gaps.push({ from: at, to });
    at = Math.max(at, span.to);
  }
  if (at < stretch.to) gaps.push({ from: at, to: stretch.to });
  return gaps;
}

/** One stretch a lock is going on, wearing the control that shuts it */
interface LockEdit extends Stretch {
  mark: Mark;
}

/** A locked cell the selection reaches, and where it stands */
interface LockedCell {
  pos: number;
  node: PMNode;
}

/**
 * The one control the stretch may be given whole, or null where it may not.
 *
 * Lifting a lock leaves the control standing, so locking a paragraph that once held one would
 * otherwise come back as three: a fresh control either side of the one already there. Taking the
 * stretch over keeps that control's id, which §17.5.2.18 asks for and dropping it would not.
 */
function takeover(
  met: readonly ControlSpan[],
  stretch: Stretch,
  open: (at: Stretch) => boolean
): LockEdit | null {
  const only = met.length === 1 ? met[0] : null;
  const prefix = only && prefixOf(only.mark.attrs);
  if (!only || isLocked(only.mark) || !prefix || !namesNothing(prefix)) {
    return null;
  }
  const mark = open(stretch) ? withLock(only.mark, true) : null;
  return mark
    ? {
        from: Math.min(stretch.from, only.from),
        to: Math.max(stretch.to, only.to),
        mark,
      }
    : null;
}

/**
 * What locking would do to one paragraph: a control of its own around each stretch of the
 * selection no control stands in, and the lock shut on each open control it covers.
 */
function blockLockEdits(
  doc: PMNode,
  block: Textblock,
  spans: readonly ControlSpan[],
  reach: Stretch
): LockEdit[] {
  const stretch = stretchIn(block, reach);
  if (!stretch) return [];
  // Inside a locked cell neither step is possible, since the guard refuses both
  const open = (at: Stretch) => !rangeTouchesLocked(doc, at.from, at.to);
  const met = spans.filter((span) => overlaps(span, stretch));
  const taken = takeover(met, stretch, open);
  if (taken) return [taken];
  const shut = met
    .filter((span) => !isLocked(span.mark) && open(span))
    .flatMap((span) => {
      const mark = withLock(span.mark, true);
      return mark ? [{ from: span.from, to: span.to, mark }] : [];
    });
  const fresh = gapsIn(stretch, spans)
    .filter(open)
    .map((gap) => ({ ...gap, mark: lockedControlMark(newControlId()) }));
  return [...shut, ...fresh];
}

/** What the selection offers to lock */
interface Lockable {
  edits: readonly LockEdit[];
}

/** What the selection reaches that is already locked */
interface Locked {
  /** The locked controls, each as the whole stretch it covers */
  spans: readonly ControlSpan[];
  cells: readonly LockedCell[];
}

/**
 * The lock state of the selection, which is the one answer the lock commands and the menus are
 * built on.
 *
 * A selection holding both something to lock and a lock to lift is `mixed`, a state of its own
 * rather than a pair of booleans a menu has to read a combination out of.
 *
 * A control the document locked against deletion alone leaves its contents open, so it reads as
 * text to lock rather than as a lock in reach: lifting that one clause by itself is not a state
 * this answer can name. The guard still refuses taking such a control away.
 */
export type SelectionLock = "none" | "lockable" | "locked" | "mixed";

/** The same answer carrying what each state would take, which is what the two commands run on */
type SelectionLockDetail =
  | { kind: "none" }
  | ({ kind: "lockable" } & Lockable)
  | ({ kind: "locked" } & Locked)
  | ({ kind: "mixed" } & Lockable & Locked);

/**
 * Reads the selection once and says what locking and unlocking would do to it.
 *
 * Reaching the edge of a control counts as reaching the control, so a caret resting against
 * either end of one unlocks it: a lock is lifted as a whole, and asking the user to select the
 * field exactly would be worse. A cell is shut as a whole in the same way, so a caret anywhere
 * inside one reaches it, and a block of selected cells reaches every locked cell in the block.
 *
 * Only a text selection has anything to lock; a whole selected image or a block of table cells
 * is not a stretch of text a control can hold.
 *
 * Settling a lock is an edit of the body, so a protection that shuts the body leaves neither
 * anything to lock nor a lock to lift (`schema/protection`).
 */
function selectionLockDetail(state: EditorState): SelectionLockDetail {
  if (editsShut(state)) return { kind: "none" };
  const selection = state.selection;
  const locking = !selection.empty && selection instanceof TextSelection;
  const edits: LockEdit[] = [];
  const spans: ControlSpan[] = [];
  // A merged cell is pointed at from every spot it covers, so each cell is counted once
  const cells = new Map<number, PMNode>();

  for (const range of selection.ranges) {
    const reach = { from: range.$from.pos, to: range.$to.pos };
    state.doc.nodesBetween(reach.from, reach.to, (node, pos) => {
      if (isLockedCell(node)) cells.set(pos, node);
      if (!node.isTextblock) return true;
      const block: Textblock = { node, start: pos + 1 };
      const controls = controlSpans(block);
      for (const span of controls) {
        if (isLocked(span.mark) && reaches(span, reach)) spans.push(span);
      }
      if (locking) {
        edits.push(...blockLockEdits(state.doc, block, controls, reach));
      }
      return false;
    });
  }

  const locked: Locked = {
    spans,
    cells: Array.from(cells, ([pos, node]) => ({ pos, node })),
  };
  const anyLocked = spans.length > 0 || locked.cells.length > 0;
  if (edits.length === 0) {
    return anyLocked ? { kind: "locked", ...locked } : { kind: "none" };
  }
  return anyLocked
    ? { kind: "mixed", edits, ...locked }
    : { kind: "lockable", edits };
}

/** What locking and unlocking would do where the selection stands */
export function selectionLock(state: EditorState): SelectionLock {
  return selectionLockDetail(state).kind;
}

/** The stretches locking would shut. Empty where the selection offers nothing to lock */
function lockEditsOf(lock: SelectionLockDetail): readonly LockEdit[] {
  switch (lock.kind) {
    case "lockable":
    case "mixed":
      return lock.edits;
    case "none":
    case "locked":
      return [];
  }
}

const NOTHING_LOCKED: Locked = { spans: [], cells: [] };

/** What lifting a lock would open. Nothing where the selection reaches no lock */
function lockedOf(lock: SelectionLockDetail): Locked {
  switch (lock.kind) {
    case "locked":
    case "mixed":
      return lock;
    case "none":
    case "lockable":
      return NOTHING_LOCKED;
  }
}

/**
 * Locks the selected text, as one content control per paragraph it runs through.
 * A control of the same type replaces the one already worn rather than being laid over it, so
 * shutting a control that already stands there is the very same step as wrapping a new one.
 * A selection that also reaches a lock still locks what that lock leaves open.
 */
export const lockSelection: Command = (state, dispatch) => {
  const edits = lockEditsOf(selectionLockDetail(state));
  if (edits.length === 0) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const edit of edits) tr.addMark(edit.from, edit.to, edit.mark);
    dispatch(tr);
  }
  return true;
};

/**
 * Lifts the lock off a cell, leaving the control that wrapped it in the file standing.
 * An opening we cannot rewrite loses its control, so that what the file says and what the editor
 * shows stay the same thing.
 */
function unlockCell(tr: Transaction, cell: LockedCell): void {
  const prefix = prefixOf(cell.node.attrs);
  tr.setNodeMarkup(cell.pos, null, {
    ...cell.node.attrs,
    sdtPrefix: prefix === null ? null : withContentLock(prefix, false),
    sdtContentsLocked: false,
    sdtDeletionLocked: false,
  });
}

/** Lifts the lock off every control the selection reaches, over that control's whole stretch */
export const unlockSelection: Command = (state, dispatch) => {
  const { spans, cells } = lockedOf(selectionLockDetail(state));
  if (spans.length === 0 && cells.length === 0) return false;
  if (dispatch) {
    const tr = state.tr.setMeta(unlockAllowed, true);
    for (const span of spans) {
      const opened = withLock(span.mark, false);
      // A control we cannot rewrite goes away instead, which beats a lock that cannot be lifted
      if (opened) tr.addMark(span.from, span.to, opened);
      else tr.removeMark(span.from, span.to, span.mark);
    }
    for (const cell of cells) unlockCell(tr, cell);
    dispatch(tr);
  }
  return true;
};

/**
 * Whether editing where the selection stands is shut.
 * This is the same question the guard asks, so a menu built on it offers nothing that would
 * then be refused: for a caret that means sitting inside a control rather than against its edge.
 *
 * A different question from `selectionLock`, deliberately: a caret against an edge lifts that
 * control's lock, and is not itself shut by it.
 */
export function selectionTouchesLocked(state: EditorState): boolean {
  return selectionShut(state.selection, state.doc);
}

/**
 * Whether the document holds any locked control at all, which is what registering a template turns
 * on.
 *
 * Either clause counts: a control whose contents may not be edited, and one that may only not be
 * taken away. This is a question about the document rather than about editing where the selection
 * stands, and a document carrying the second kind is a document with a lock in it.
 */
export function documentHasLocked(doc: PMNode): boolean {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    if (carriesLock(node)) found = true;
    return !found;
  });
  return found;
}
