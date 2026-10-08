import { Mark, type Node as PMNode, type ResolvedPos } from "prosemirror-model";
import { ownBackground } from "../docx/runProps";
import { docxSchema } from "../schema";
import { isEmptyRun } from "../schema/emptyRuns";
import { drawsNothing, type StepRange } from "../schema/locks";
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

/** Whether the node shows on the page, a run holding no characters drawn as a blank among them */
function shows(node: PMNode): boolean {
  if (!drawsNothing(node)) return true;
  const run = isEmptyRun(node) ? runOf(node) : null;
  return run !== null && ownBackground(run) !== null;
}

/** The index of the nearest child that matches, looking before the caret first. null for none */
function nearest(
  parent: PMNode,
  index: number,
  matches: (node: PMNode) => boolean
): number | null {
  for (let at = index - 1; at >= 0; at -= 1) {
    if (matches(parent.child(at))) return at;
  }
  for (let at = index; at < parent.childCount; at += 1) {
    if (matches(parent.child(at))) return at;
  }
  return null;
}

/** The run holding no characters at this index, and those of its kind beside it inside the same wrappers */
function filledFrom($pos: ResolvedPos, index: number): StepRange {
  const parent = $pos.parent;
  const wrappers = wrapperMarks(parent.child(index));
  const joins = (at: number) => {
    const node = parent.maybeChild(at);
    return (
      node !== null &&
      isEmptyRun(node) &&
      Mark.sameSet(wrapperMarks(node), wrappers)
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
 */
export function formatSourceAt(doc: PMNode, pos: number): FormatSource | null {
  const $pos = doc.resolve(pos);
  const parent = $pos.parent;
  if (!parent.inlineContent) return null;
  const index = $pos.index();
  if ($pos.textOffset > 0) {
    return { run: runOf(parent.child(index)), fills: null };
  }
  const source =
    nearest(parent, index, shows) ?? nearest(parent, index, isEmptyRun);
  if (source === null) return null;
  const node = parent.child(source);
  return {
    run: runOf(node),
    fills: isEmptyRun(node) ? filledFrom($pos, source) : null,
  };
}
