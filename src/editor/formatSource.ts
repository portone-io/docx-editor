import { Mark, type Node as PMNode, type ResolvedPos } from "prosemirror-model";
import { ownBackground } from "../docx/runProps";
import { docxSchema } from "../schema";
import { isEmptyRun } from "../schema/emptyRuns";
import {
  controlsWrittenInto,
  drawsNothing,
  type StepRange,
} from "../schema/locks";
import { wrapperMarks } from "../schema/wrappers";

export interface FormatSource {
  /** The run mark what is written takes, null for none */
  run: Mark | null;
  /** The runs holding no characters what is written fills, where one of them gave the formatting */
  fills: StepRange | null;
}

function runOf(node: PMNode): Mark | null {
  return docxSchema.marks.run.isInSet(node.marks) ?? null;
}

export function sameRun(a: Mark | null, b: Mark | null): boolean {
  return a === null || b === null ? a === b : a.eq(b);
}

/**
 * Whether text written beside this run holding no characters can take its place.
 *
 * Written text never takes a wrapper on, so filling the run would take a wrapper that stood on
 * nothing else away with it. Only a control can be laid over the text again, where it is open; any
 * other wrapper keeps its run.
 */
function fillable(doc: PMNode, pos: number, node: PMNode): boolean {
  if (!isEmptyRun(node)) return false;
  const open = controlsWrittenInto(doc, pos, pos + node.nodeSize).map(
    (span) => span.mark
  );
  return wrapperMarks(node).every(
    (wrapper) =>
      wrapper.type === docxSchema.marks.sdt &&
      open.some((control) => control.eq(wrapper))
  );
}

/**
 * Whether the node shows a character text written beside it takes the formatting of, a run holding
 * no characters drawn as a blank among them where the text fills it
 */
function shows(doc: PMNode, node: PMNode, pos: number): boolean {
  if (!drawsNothing(node)) return true;
  const run = runOf(node);
  return (
    run !== null && ownBackground(run) !== null && fillable(doc, pos, node)
  );
}

/** The index of the nearest child that matches, looking before the caret first. null for none */
function nearest(
  $pos: ResolvedPos,
  matches: (node: PMNode, pos: number) => boolean
): number | null {
  const parent = $pos.parent;
  const index = $pos.index();
  const at = (child: number) =>
    matches(parent.child(child), $pos.posAtIndex(child));
  for (let child = index - 1; child >= 0; child -= 1) {
    if (at(child)) return child;
  }
  for (let child = index; child < parent.childCount; child += 1) {
    if (at(child)) return child;
  }
  return null;
}

/** The run holding no characters at this index, and those beside it wearing its marks */
function filledFrom($pos: ResolvedPos, index: number): StepRange {
  const parent = $pos.parent;
  const source = parent.child(index);
  const joins = (at: number) => {
    const node = parent.maybeChild(at);
    return (
      node !== null &&
      isEmptyRun(node) &&
      sameRun(runOf(node), runOf(source)) &&
      Mark.sameSet(wrapperMarks(node), wrapperMarks(source))
    );
  };
  let first = index;
  while (joins(first - 1)) first -= 1;
  let last = index;
  while (joins(last + 1)) last += 1;
  return { from: $pos.posAtIndex(first), to: $pos.posAtIndex(last + 1) };
}

/**
 * Word's rule: the nearest character the paragraph shows, before the caret else after it, and a run
 * holding no characters only where the paragraph shows none. null in a paragraph holding nothing.
 *
 * A run holding no characters gives the formatting only where what is written fills it, so a blank
 * the text cannot take the place of is passed over and never stands painted twice.
 */
export function formatSourceAt(doc: PMNode, pos: number): FormatSource | null {
  const $pos = doc.resolve(pos);
  const parent = $pos.parent;
  if (!parent.inlineContent) return null;
  if ($pos.textOffset > 0) {
    return { run: runOf(parent.child($pos.index())), fills: null };
  }
  const source =
    nearest($pos, (node, at) => shows(doc, node, at)) ??
    nearest($pos, (node, at) => fillable(doc, at, node));
  if (source === null) return null;
  const node = parent.child(source);
  return {
    run: runOf(node),
    fills: isEmptyRun(node) ? filledFrom($pos, source) : null,
  };
}
