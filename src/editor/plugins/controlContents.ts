/**
 * Keeps an inline content control standing through an edit of what it holds
 * (`spec/notes/contentControls.md`).
 *
 * An inline control is a mark, and the mark is not inclusive, so ProseMirror carries it over a
 * replacement only where it stands on both sides of it. Nothing of the control stands past its
 * end, so typing, a composition or a paste over everything the control holds, or over its end,
 * would write the new text beside the control, and a deletion of everything it holds would take
 * the control away with its text.
 *
 * The lock guard reads such a replacement as an edit of the contents (`schema/locks`), and this
 * appends what makes it one: the control laid back over what the edit wrote or, where the edit
 * wrote nothing, the control left standing as one holding nothing (`docx/wrappers`).
 *
 * The caret such an edit leaves at the end of the control stays inside it, and so does the one
 * left beside a control the edit emptied: what is typed, composed or pasted there next goes into
 * the control rather than beside it, until the caret moves. A caret the user places against the
 * edge of a control still stands outside it.
 *
 * It also owns where what is written at a caret takes its formatting from (`../formatSource`).
 * A caret beside something that draws nothing writes with the formatting of the character the rule
 * finds, and what is written where a run holding no characters gave the formatting takes that
 * run's formatting and fills it, a paste included: the run goes, and where it stood inside open
 * controls, the controls are laid over what was written, so a caret placed in a control's blank
 * writes into the control. A blank standing in any other wrapper, a control whose lock shuts its
 * contents or a link among them, stays, and the text stands beside it.
 */

import { Mark, type Node as PMNode, type Slice } from "prosemirror-model";
import {
  Plugin,
  PluginKey,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import {
  Mapping,
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
  type StepMap,
} from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";
import { docxSchema } from "../../schema";
import {
  controlAttrs,
  controlFactsOf,
  goesWithEdit,
  isEmptyInlineControl,
  OWN_CONTROL_ATTRS,
} from "../../schema/controlAttrs";
import { displayOnly } from "../../schema/displayDerivation";
import { isEmptyRun } from "../../schema/emptyRuns";
import {
  controlKept,
  controlLifted,
  controlsWrittenInto,
  historyReplay,
  type StepRange,
} from "../../schema/locks";
import {
  innermostDepth,
  isWrapperType,
  wrapperMarks,
  wrappersOf,
} from "../../schema/wrappers";
import { formatSourceAt } from "../formatSource";

/** Where the caret keeps writing into the controls it wrote into, while it stays there */
type Continuing = number | null;

const controlContentsKey = new PluginKey<Continuing>(
  "docxEditorControlContents"
);

/** What this plugin hands its own state, told apart from a transaction saying nothing */
interface ContinuingMeta {
  continuing: Continuing;
}

/**
 * The passes that say a change is no edit of the contents. A history replay puts back a state that
 * was already settled, a re-derivation writes no content, and the other two are what an edit does
 * to a control rather than an edit of their own.
 */
const NOT_AN_EDIT = [historyReplay, displayOnly, controlLifted, controlKept];

function userEdit(tr: Transaction): boolean {
  return (
    tr.docChanged &&
    tr.getMeta("appendedTransaction") === undefined &&
    !NOT_AN_EDIT.some((pass) => tr.getMeta(pass) === true)
  );
}

/** Every content control this inline node wears, outermost first */
function controlsOf(node: PMNode | null | undefined): readonly Mark[] {
  return node ? wrappersOf(node, docxSchema.marks.sdt) : [];
}

/** The control a node holding nothing stands for, as the mark it wears once it holds text again */
function controlOfEmpty(node: PMNode): Mark {
  return docxSchema.marks.sdt.create({
    depth: innermostDepth(node.marks) + 1,
    key: node.attrs.key,
    ...controlAttrs(
      OWN_CONTROL_ATTRS,
      controlFactsOf(OWN_CONTROL_ATTRS, node.attrs)
    ),
  });
}

/** The control this mark stands for, as the node it stands as once it holds nothing */
function emptyOf(control: Mark, wearing: readonly Mark[]): PMNode {
  return docxSchema.nodes.sdtEmptyInline.create(
    {
      key: control.attrs.key,
      ...controlAttrs(
        OWN_CONTROL_ATTRS,
        controlFactsOf(OWN_CONTROL_ATTRS, control.attrs)
      ),
    },
    null,
    wearing
  );
}

/**
 * Whether two carriers stand for one and the same control: a mark and the node it stands as once
 * it holds nothing, or two marks that differ in nothing but the depth a node holding nothing
 * could not keep.
 */
function sameControl(a: Pick<Mark, "attrs">, b: Pick<Mark, "attrs">): boolean {
  return a.attrs.key === b.attrs.key && a.attrs.sdtPrefix === b.attrs.sdtPrefix;
}

function wearsControl(node: PMNode | null | undefined, control: Mark): boolean {
  return controlsOf(node).some((worn) => sameControl(worn, control));
}

/** Whether the control is kept at all, rather than going with the edit (`w:temporary`) */
function kept(control: Mark): boolean {
  return !goesWithEdit(controlFactsOf(OWN_CONTROL_ATTRS, control.attrs));
}

/**
 * The controls a caret at this spot writes into when it continues: those ending here, or the one
 * standing here holding nothing.
 */
function controlsContinued(doc: PMNode, pos: number): readonly Mark[] {
  const $pos = doc.resolve(pos);
  const before = $pos.nodeBefore;
  if (!before) return [];
  if (isEmptyInlineControl(before)) {
    const control = controlOfEmpty(before);
    return kept(control) ? [control] : [];
  }
  const after = controlsOf($pos.nodeAfter);
  return controlsOf(before).filter(
    (control) => kept(control) && !after.some((other) => other.eq(control))
  );
}

/** One edit that wrote into controls, as the stretch it wrote in the document that came out */
interface Written {
  /** The controls it wrote into, outermost first */
  controls: readonly Mark[];
  from: number;
  to: number;
  /** The node the controls began at, whose formatting and wrappers the controls' text wore */
  first: PMNode | null;
}

/** What was written where a run holding no characters gave the formatting, in the document that came out */
interface Fill {
  /** The runs holding no characters it filled, which go */
  runs: StepRange;
  /** What was written */
  written: StepRange;
  /** The run mark the runs gave, which it takes where `restyles` */
  run: Mark | null;
  /** Whether it takes that mark, rather than keeping the marks the caret was given on purpose */
  restyles: boolean;
}

/** What these transactions wrote, in the document that came out */
interface Writes {
  /** The stretches written into controls */
  written: Written[];
  filled: Fill[];
}

/** What a step puts in, at the stretch it puts it in place of */
interface Insertion {
  from: number;
  to: number;
  slice: Slice;
}

function insertionOf(step: Step): Insertion | null {
  if (step instanceof ReplaceStep) {
    return { from: step.from, to: step.to, slice: step.slice };
  }
  // A paste whose last paragraph stands deeper than the one it lands in carries the text after
  // the caret into that paragraph around a gap, and puts its own in front of it
  if (step instanceof ReplaceAroundStep) {
    return { from: step.from, to: step.gapFrom, slice: step.slice };
  }
  return null;
}

/**
 * Every stretch these transactions wrote into a control, and every run holding no characters they
 * filled, in the document that came out.
 *
 * A step writes into a control when it replaces a stretch the control holds, up to all of it
 * (`controlsWrittenInto`), when it puts content in at the spot a caret continuing inside a
 * control stood, or when it fills a run holding no characters that stands inside the control.
 */
function writtenStretches(
  transactions: readonly Transaction[],
  continuing: Continuing,
  storedBefore: readonly Mark[] | null
): Writes {
  const maps: StepMap[] = transactions.flatMap((tr) => tr.mapping.maps);
  const writes: Writes = { written: [], filled: [] };
  let index = 0;
  let stored = storedBefore;
  for (const tr of transactions) {
    const edit = userEdit(tr);
    // Text typed with marks the caret was given was formatted on purpose; a paste carries its own
    const restyles = stored === null || pasted(tr);
    stored = tr.storedMarks;
    tr.steps.forEach((step, stepIndex) => {
      const at = index;
      index += 1;
      const insertion = edit ? insertionOf(step) : null;
      if (insertion === null) return;
      const doc = tr.docs[stepIndex] ?? tr.doc;
      const rest = new Mapping(maps.slice(at + 1));
      const from = rest.map(insertion.from, 1);
      const writtenTo = Math.max(
        from,
        rest.map(insertion.from + insertion.slice.size, -1)
      );
      const fill = stepFill(doc, insertion);
      if (fill !== null) {
        const through = new Mapping(maps.slice(at));
        writes.filled.push({
          runs: {
            from: through.map(fill.runs.from, 1),
            to: through.map(fill.runs.to, -1),
          },
          written: { from, to: writtenTo },
          run: fill.run,
          restyles,
        });
      }
      const found =
        step instanceof ReplaceStep
          ? stepControls(doc, step, continuing, maps.slice(0, at))
          : null;
      const into = found ?? fill;
      if (into === null || into.controls.length === 0) return;
      writes.written.push({
        controls: into.controls,
        first: into.first,
        from,
        to: writtenTo,
      });
    });
  }
  return writes;
}

/**
 * Whether the slice puts text in. An Enter splitting a paragraph puts in none, and neither does a
 * comment's markers being laid around a stretch, which write nothing in the run beside them.
 */
function writesText(slice: Slice): boolean {
  let found = false;
  slice.content.descendants((node) => {
    if (node.isText) found = true;
    return !found;
  });
  return found;
}

/** Whether the transaction is a paste or a drop, which carries the marks of its own */
function pasted(tr: Transaction): boolean {
  const event: unknown = tr.getMeta("uiEvent");
  return event === "paste" || event === "drop";
}

/**
 * The runs holding no characters an insertion at a caret fills, the formatting they give, the
 * controls around them it writes into, and the first of the runs. null where it fills none.
 *
 * Inserted text never takes a wrapper on, so filling a run would take a wrapper that stood on
 * nothing else away with it. Only a control can be laid over the text again, where it is open; one
 * going with any edit (`w:temporary`) goes with the run. Any other wrapper keeps its run, and the
 * text stands beside it.
 */
function stepFill(
  doc: PMNode,
  insertion: Insertion
):
  | (Pick<Written, "controls" | "first"> & {
      runs: StepRange;
      run: Mark | null;
    })
  | null {
  if (insertion.from !== insertion.to || !writesText(insertion.slice)) {
    return null;
  }
  const source = formatSourceAt(doc, insertion.from);
  const runs = source?.fills ?? null;
  const first = runs === null ? null : doc.nodeAt(runs.from);
  if (source === null || runs === null || first === null) return null;
  const open = controlsWrittenInto(doc, runs.from, runs.to).map(
    (span) => span.mark
  );
  const laidAgain = wrapperMarks(first).every(
    (wrapper) =>
      wrapper.type === docxSchema.marks.sdt &&
      open.some((control) => control.eq(wrapper))
  );
  if (!laidAgain) return null;
  return {
    runs,
    run: source.run,
    controls: controlsOf(first).filter(kept),
    first,
  };
}

/** The controls one step writes into, and what the content it replaced wore. null for none */
function stepControls(
  doc: PMNode,
  step: ReplaceStep,
  continuing: Continuing,
  before: readonly StepMap[]
): Pick<Written, "controls" | "first"> | null {
  if (step.from < step.to) {
    const spans = controlsWrittenInto(doc, step.from, step.to);
    const controls = spans.map((span) => span.mark).filter(kept);
    if (controls.length === 0) return null;
    const innermost = Math.max(...spans.map((span) => span.from));
    return { controls, first: doc.nodeAt(innermost) };
  }
  if (continuing === null) return null;
  if (new Mapping(before).map(continuing) !== step.from) return null;
  const controls = controlsContinued(doc, step.from);
  if (controls.length === 0) return null;
  return { controls, first: doc.resolve(step.from).nodeBefore };
}

/** Whether this node carries the wrapper, a control standing as the node itself included */
function carries(node: PMNode, wrapper: Mark): boolean {
  if (wrapper.type !== docxSchema.marks.sdt) return wrapper.isInSet(node.marks);
  return isEmptyInlineControl(node)
    ? sameControl(node, wrapper)
    : wearsControl(node, wrapper);
}

/** Whether a wrapper still stands anywhere in the document */
function stands(doc: PMNode, wrapper: Mark): boolean {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    found = carries(node, wrapper);
    return !found;
  });
  return found;
}

/**
 * The marks a caret continuing after these controls writes with: the formatting of what they held
 * and the wrappers around it, the controls among them.
 */
function marksWith(
  worn: readonly Mark[],
  controls: readonly Mark[]
): readonly Mark[] {
  let set: readonly Mark[] = Mark.none;
  for (const mark of [...worn, ...controls]) set = mark.addToSet(set);
  return set;
}

/**
 * Lays each control back over what the edit wrote into it, and leaves one that holds nothing any
 * more standing as a node where its text stood. Answers the marks a caret beside such a node should
 * write with, which are what the text there wore, or null where no caret stands beside one.
 */
function keepControls(
  tr: Transaction,
  written: readonly Written[]
): readonly Mark[] | null {
  let stored: readonly Mark[] | null = null;
  for (const stretch of written) {
    const from = tr.mapping.map(stretch.from, 1);
    const $from = tr.doc.resolve(from);
    if (!$from.parent.inlineContent) continue;
    // What a paste wrote past the paragraph the control stands in cannot be inside it
    const to = Math.min(tr.mapping.map(stretch.to, -1), $from.end());
    if (to > from) {
      for (const control of stretch.controls) {
        if (!wearsControl(tr.doc.nodeAt(from), control)) {
          tr.addMark(from, to, control);
        }
      }
    }
    const gone = stretch.controls.filter((control) => !stands(tr.doc, control));
    const innermost = gone.at(-1);
    if (innermost === undefined) continue;
    const node = emptyOf(innermost, wrappersAround(tr.doc, stretch, gone));
    tr.insert(from, node);
    const { selection } = tr;
    if (selection.empty && selection.head === from + node.nodeSize) {
      stored = marksWith(
        (stretch.first?.marks ?? []).filter(
          (mark) => !isWrapperType(mark.type)
        ),
        [...node.marks, innermost]
      );
    }
  }
  return stored;
}

/**
 * The wrappers a control left holding nothing stands inside: those the text it held stood inside,
 * outside the control, that are still there - another control emptied with it among them.
 */
function wrappersAround(
  doc: PMNode,
  stretch: Written,
  gone: readonly Mark[]
): readonly Mark[] {
  const ordered = stretch.first ? wrapperMarks(stretch.first) : [];
  const innermost = gone.at(-1);
  const inside = ordered.findIndex(
    (mark) =>
      innermost !== undefined &&
      mark.type === innermost.type &&
      sameControl(mark, innermost)
  );
  return ordered
    .slice(0, inside < 0 ? 0 : inside)
    .filter(
      (mark) =>
        gone.some(
          (control) => mark.type === control.type && sameControl(control, mark)
        ) || stands(doc, mark)
    );
}

/**
 * Takes away each node standing for a control that holds text again beside it, which is what
 * typing into a control the caret emptied leaves: the text wears the control, and the node that
 * stood for it while it held nothing would be a second copy of it.
 */
function dropRefilled(tr: Transaction, written: readonly Written[]): void {
  const blocks = new Set<number>();
  for (const stretch of written) {
    const $at = tr.doc.resolve(tr.mapping.map(stretch.from, 1));
    if ($at.parent.inlineContent) blocks.add($at.before());
  }
  const refilled: { from: number; to: number }[] = [];
  for (const before of blocks) {
    const block = tr.doc.nodeAt(before);
    if (!block) continue;
    block.forEach((child, offset, index) => {
      if (!isEmptyInlineControl(child)) return;
      const beside = [block.maybeChild(index - 1), block.maybeChild(index + 1)];
      const refills = beside.some((node) =>
        controlsOf(node).some((control) => sameControl(child, control))
      );
      const from = before + 1 + offset;
      if (refills) refilled.push({ from, to: from + child.nodeSize });
    });
  }
  for (const { from, to } of refilled.reverse()) tr.delete(from, to);
}

function sameRun(a: Mark | null, b: Mark | null): boolean {
  return a === null || b === null ? a === b : a.eq(b);
}

/** Gives what was written into runs holding no characters the formatting they gave */
function restyleFilled(tr: Transaction, filled: readonly Fill[]): void {
  const runType = docxSchema.marks.run;
  for (const { written, run, restyles } of filled) {
    if (!restyles) continue;
    const from = tr.mapping.map(written.from, 1);
    const $from = tr.doc.resolve(from);
    if (!$from.parent.inlineContent) continue;
    // What a paste wrote past the paragraph the runs stood in took no formatting from them
    const to = Math.min(tr.mapping.map(written.to, -1), $from.end());
    const differing: StepRange[] = [];
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (
        !node.isText &&
        !(node.isInline && node.type.allowsMarkType(runType))
      ) {
        return;
      }
      if (sameRun(runType.isInSet(node.marks) ?? null, run)) return;
      differing.push({
        from: Math.max(pos, from),
        to: Math.min(pos + node.nodeSize, to),
      });
    });
    for (const range of differing) {
      tr.removeMark(range.from, range.to, runType);
      if (run !== null) tr.addMark(range.from, range.to, run);
    }
  }
}

/** Takes away the runs holding no characters that what was written filled */
function dropFilled(tr: Transaction, filled: readonly Fill[]): void {
  // Two steps of one edit may fill the same runs, which go once
  const runs = new Map<number, number>();
  for (const { runs: stretch } of filled) {
    const from = tr.mapping.map(stretch.from, 1);
    const to = tr.mapping.map(stretch.to, -1);
    if (from >= to) continue;
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (isEmptyRun(node)) runs.set(pos, pos + node.nodeSize);
      return node.isBlock;
    });
  }
  const lastFirst = [...runs].sort(([a], [b]) => b - a);
  for (const [from, to] of lastFirst) tr.delete(from, to);
}

/**
 * Where the caret keeps writing into a control after this change: where an edit that wrote into a
 * control left it, or where it already continued, as long as a control still ends there.
 */
/**
 * The marks the caret writes with where the formatting rule (`../formatSource`) parts from those
 * ProseMirror reads off the node beside it. null where the two agree.
 */
function caretMarks(tr: Transaction): readonly Mark[] | null {
  const caret =
    tr.selection instanceof TextSelection ? tr.selection.$cursor : null;
  const source = caret ? formatSourceAt(tr.doc, caret.pos) : null;
  if (!caret || source === null) return null;
  const own = caret.marks();
  const runType = docxSchema.marks.run;
  if (sameRun(runType.isInSet(own) ?? null, source.run)) return null;
  const others = runType.removeFromSet(own);
  return source.run === null ? others : source.run.addToSet(others);
}

function continuingAfter(
  tr: Transaction,
  written: readonly Written[],
  previous: Continuing
): Continuing {
  const caret =
    tr.selection instanceof TextSelection ? tr.selection.$cursor : null;
  if (!caret || controlsContinued(tr.doc, caret.pos).length === 0) {
    return null;
  }
  const left =
    written.some((stretch) => tr.mapping.map(stretch.to, 1) === caret.pos) ||
    (previous !== null && tr.mapping.map(previous) === caret.pos);
  return left ? caret.pos : null;
}

export function controlContents(): Plugin<Continuing> {
  let live: EditorView | null = null;
  return new Plugin<Continuing>({
    key: controlContentsKey,
    state: {
      init: () => null,
      apply(tr, value, oldState, newState): Continuing {
        const meta: ContinuingMeta | undefined = tr.getMeta(controlContentsKey);
        if (meta !== undefined) return meta.continuing;
        if (value === null || tr.getMeta(historyReplay) === true) return null;
        if (tr.docChanged) return tr.mapping.map(value);
        return newState.selection.eq(oldState.selection) ? value : null;
      },
    },
    view(view) {
      live = view;
      return {
        destroy() {
          live = null;
        },
      };
    },
    appendTransaction(transactions, oldState, newState) {
      const edited = transactions.some(userEdit);
      if (!edited && !transactions.some((tr) => tr.selectionSet)) return null;
      const { written, filled }: Writes = edited
        ? writtenStretches(
            transactions,
            controlContentsKey.getState(oldState) ?? null,
            oldState.storedMarks
          )
        : { written: [], filled: [] };
      const previous = controlContentsKey.getState(newState) ?? null;
      const tr = newState.tr.setMeta(controlKept, true);
      // The controls go on before the runs come off, so a control never stands on nothing in
      // between those two steps
      const stored = keepControls(tr, written);
      dropRefilled(tr, written);
      restyleFilled(tr, filled);
      dropFilled(tr, filled);
      // Stored marks are how ProseMirror ends a composition, so none are set under an open one
      if (live?.composing !== true) {
        const marks =
          stored ?? (tr.storedMarks === null ? caretMarks(tr) : null);
        if (marks !== null) tr.setStoredMarks(marks);
      }
      const continuing = continuingAfter(tr, written, previous);
      if (!tr.docChanged && !tr.storedMarksSet && continuing === previous) {
        return null;
      }
      const meta: ContinuingMeta = { continuing };
      return tr.setMeta(controlContentsKey, meta);
    },
  });
}
