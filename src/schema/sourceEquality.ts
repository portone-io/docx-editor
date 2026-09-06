/**
 * Whether two nodes would be written back as the same XML.
 *
 * `Node.eq` answers a different question: whether they are the same node, display values and all.
 * That is the right question inside an editing session, where both sides came from one import and
 * a difference can only be an edit. It is the wrong one across a re-derivation, because the values
 * worked out from the surroundings (`./attrRoles`) move without the document moving, and a block
 * judged changed is a block rewritten, which costs it the markup the writer does not model.
 *
 * So the comparison is `Node.eq` with the display attrs left out, and nothing else left out: the
 * type, the text, the marks and the source and session attrs all have to agree, and the recursion
 * runs the same way, so a node that says the same thing in a different shape is still a change.
 */

import {
  Fragment,
  type Mark,
  type MarkType,
  type NodeType,
  type Node as PMNode,
} from "prosemirror-model";
import { displayAttrsOf } from "./attrRoles";

/** Whether two attr values are the same, reading through the arrays and objects attrs are allowed to hold */
function sameValue(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a) || Array.isArray(b)) {
    return (
      Array.isArray(a) &&
      Array.isArray(b) &&
      a.length === b.length &&
      a.every((item, at) => sameValue(item, b[at]))
    );
  }
  if (
    typeof a !== "object" ||
    typeof b !== "object" ||
    a === null ||
    b === null
  ) {
    return false;
  }
  const left = Object.entries(a);
  return (
    left.length === Object.keys(b).length &&
    left.every(
      ([key, value]) =>
        key in b &&
        sameValue(value, (b as Readonly<Record<string, unknown>>)[key])
    )
  );
}

function sameAttrs(
  display: readonly string[],
  a: Readonly<Record<string, unknown>>,
  b: Readonly<Record<string, unknown>>
): boolean {
  return Object.keys(a)
    .filter((name) => !display.includes(name))
    .every((name) => sameValue(a[name], b[name]));
}

function sameMarks(a: readonly Mark[], b: readonly Mark[]): boolean {
  return (
    a.length === b.length &&
    a.every((mark, at) => {
      const other = b[at];
      return (
        other !== undefined &&
        mark.type === other.type &&
        sameAttrs(displayAttrsOf(mark.type), mark.attrs, other.attrs)
      );
    })
  );
}

/**
 * Whether the two would write the same XML: the same type, text, marks and source attrs, all the
 * way down. What the editor works out for the screen is not compared.
 */
export function sameSource(a: PMNode, b: PMNode): boolean {
  if (a === b) return true;
  if (
    a.type !== b.type ||
    !sameMarks(a.marks, b.marks) ||
    !sameAttrs(displayAttrsOf(a.type), a.attrs, b.attrs)
  ) {
    return false;
  }
  if (a.isText) return a.text === b.text;
  if (a.childCount !== b.childCount) return false;
  for (let at = 0; at < a.childCount; at += 1) {
    if (!sameSource(a.child(at), b.child(at))) return false;
  }
  return true;
}

/** The attrs with every worked-out one back at the value the schema starts it on */
function withoutDisplay(
  type: MarkType | NodeType,
  attrs: Readonly<Record<string, unknown>>
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(attrs).map(([name, value]) =>
      displayAttrsOf(type).includes(name)
        ? [name, type.spec.attrs?.[name]?.default ?? null]
        : [name, value]
    )
  );
}

function withoutDisplayMarkAttrs(mark: Mark): Mark {
  return displayAttrsOf(mark.type).length === 0
    ? mark
    : mark.type.create(withoutDisplay(mark.type, mark.attrs));
}

const childrenOf = (node: PMNode): PMNode[] =>
  Array.from({ length: node.childCount }, (_, at) => node.child(at));

/**
 * The node with every worked-out attr back at the value the schema starts it on, on itself and on
 * everything inside it.
 *
 * `sameSource` needs no such copy, and takes none. This is for a caller that has to hand a whole
 * document to something comparing the ordinary way.
 */
export function withoutDisplayAttrs(node: PMNode): PMNode {
  const marks = node.marks.map(withoutDisplayMarkAttrs);
  if (node.isText) return node.mark(marks);
  return node.type.create(
    withoutDisplay(node.type, node.attrs),
    Fragment.fromArray(
      childrenOf(node).map((child) => withoutDisplayAttrs(child))
    ),
    marks
  );
}
