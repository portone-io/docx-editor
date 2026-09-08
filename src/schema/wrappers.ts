/**
 * The inline wrappers a stretch of content stands inside, as the schema records them.
 *
 * A wrapper is an element that holds inline content and goes back out around the same content:
 * a content control (`w:sdt`), a hyperlink (`w:hyperlink`). Each is a mark of its own, so each
 * keeps the `inclusive` rule and the attrs its kind needs, and all of them share the `wrapper`
 * group, so a node's `marks:` whitelist names the group rather than every kind in it and a new
 * kind arrives as one mark spec.
 *
 * The nesting is recorded on the marks themselves: `depth` counts from the outermost wrapper the
 * file wrote inwards. The order of a mark set cannot carry it - `Mark.addToSet` keeps insertion
 * order only among marks of different types, `parseDOM` rebuilds a set from the page, and two
 * wrappers of one kind may stand one inside the other - so `depth` is the source of truth and
 * `wrapperMarks` is the one place that reads the order off it.
 *
 * What each kind reads out of the file and writes back into it is `docx/wrappers`.
 */

import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";

/** The mark group every inline wrapper belongs to */
export const WRAPPER_GROUP = "wrapper";

/** What every wrapper mark spec carries besides the attrs of its own kind */
export const WRAPPER_ATTRS = {
  /** Outer-to-inner order as read from the file; equal depths fall back to declaration rank */
  depth: { default: 0 },
  /** Which occurrence this is in the opened document, so two wrappers written alike stay apart */
  key: { default: 0 },
} as const;

/**
 * What the schema says about one mark type, worked out once. Both answers are read for every
 * inline node the writer and the guards walk, and neither can change over a type's life.
 */
interface TypeFacts {
  /** Whether marks of this type wrap the content they cover (`WRAPPER_GROUP`) */
  wrapper: boolean;
  /** Where the type stands among the schema's marks, which is the order they were declared in */
  rank: number;
}

const FACTS = new WeakMap<MarkType, TypeFacts>();

function factsOf(type: MarkType): TypeFacts {
  const known = FACTS.get(type);
  if (known) return known;
  const facts: TypeFacts = {
    wrapper: (type.spec.group ?? "").split(" ").includes(WRAPPER_GROUP),
    rank: Object.keys(type.schema.marks).indexOf(type.name),
  };
  FACTS.set(type, facts);
  return facts;
}

/** Whether marks of this type wrap the content they cover (`WRAPPER_GROUP`) */
export function isWrapperType(type: MarkType): boolean {
  return factsOf(type).wrapper;
}

function depthOf(mark: Mark): number {
  const depth: unknown = mark.attrs.depth;
  return typeof depth === "number" ? depth : 0;
}

/**
 * The wrappers this node stands inside, outermost first.
 *
 * Two wrappers written at the same depth - a control laid over a link during editing, say - fall
 * back to the order the schema declares them in, which is the nesting the marks meant before
 * `depth` existed.
 */
export function wrapperMarks(node: PMNode): readonly Mark[] {
  const marks = node.marks;
  if (marks.length === 0) return marks;
  const wrappers = marks.filter((mark) => factsOf(mark.type).wrapper);
  if (wrappers.length < 2) return wrappers;
  return wrappers.sort(
    (a, b) =>
      depthOf(a) - depthOf(b) || factsOf(a.type).rank - factsOf(b.type).rank
  );
}

/** Every wrapper of this kind the node stands inside, outermost first */
export function wrappersOf(node: PMNode, name: string): readonly Mark[] {
  return wrapperMarks(node).filter((mark) => mark.type.name === name);
}

/** The outermost wrapper of this kind the node stands inside. null where it stands in none */
export function wrapperOf(node: PMNode, name: string): Mark | null {
  return wrapperMarks(node).find((mark) => mark.type.name === name) ?? null;
}

/**
 * The depth of the innermost of these wrappers, and -1 where there is none, so that
 * `innermostDepth(marks) + 1` is the depth a wrapper laid inside every one of them takes.
 */
export function innermostDepth(marks: readonly Mark[]): number {
  return marks.reduce(
    (deepest, mark) =>
      factsOf(mark.type).wrapper ? Math.max(deepest, depthOf(mark)) : deepest,
    -1
  );
}

/**
 * The wrappers all of these nodes stand inside alike, outermost first: the longest run of
 * wrappers, from the outside in, that every one of them wears the very same one of.
 *
 * A wrapper laid across the whole stretch goes inside these and outside everything else, which is
 * what keeps it one element in the file rather than one per stretch the nesting differs over.
 */
export function sharedWrappers(nodes: readonly PMNode[]): readonly Mark[] {
  const first = nodes.at(0);
  if (first === undefined) return [];
  let shared = wrapperMarks(first);
  for (const node of nodes.slice(1)) {
    const marks = wrapperMarks(node);
    const same = shared.findIndex(
      (mark, index) => !(marks[index] && mark.eq(marks[index]))
    );
    if (same !== -1) shared = shared.slice(0, same);
    if (shared.length === 0) break;
  }
  return shared;
}
