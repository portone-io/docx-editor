/**
 * The guards over the nodes a document is opened with and the editor only preserves: the two ends
 * of a bookmark range, and the reference standing where a footnote or an endnote is called.
 *
 * Neither is content the editor writes, so nothing in it can put such a node back once it is gone,
 * and a document that lost one of a bookmark's two ends cannot be written back as a file at all
 * (`docx/exportDocx` refuses it). What each guard holds is therefore the whole list of the nodes it
 * answers for, in the order the document carries them.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditGuard, rangeHolds, transactionReaches } from "./editGuard";

/** Everything about one preserved node that has to read the same after a change as before it */
type Signature = (node: PMNode) => string;

function signatures(
  doc: PMNode,
  holds: (node: PMNode) => boolean,
  signature: Signature
): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    if (holds(node)) found.push(signature(node));
    return true;
  });
  return found;
}

function same(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

/**
 * A guard that lets a change through only while it leaves the list of these nodes reading exactly
 * as it did, which holds their number, their contents and their order at once.
 *
 * The comparison walks the whole document, so it is reached for only once a step of the change has
 * touched such a node: ordinary typing costs the stretches its own steps rewrote and nothing more.
 * It is judged over the change rather than step by step because a marker moved whole is two steps,
 * one taking it out and one putting it back, and neither of them alone leaves the list as it was.
 *
 * Nothing lifts it. A pass is how an edit reaches past a lock the document put on itself, and
 * there is no such thing to reach past here: the file simply has no way to say what a lost marker
 * was.
 */
export function preservedNodeGuard(
  name: string,
  holds: (node: PMNode) => boolean,
  signature: Signature
): EditGuard {
  const listOf = (doc: PMNode) => signatures(doc, holds, signature);
  return {
    name,
    change: (tr) =>
      !transactionReaches(tr, holds) || same(listOf(tr.before), listOf(tr.doc)),
    shuts: (intent, state) =>
      intent.kind === "replace" &&
      rangeHolds(state.doc, intent.from, intent.to, holds),
  };
}

/** A bookmark marker inside a paragraph arrives as raw XML, so it is known by what that XML reads */
const BOOKMARK_XML = /<(?:[\w.-]+:)?bookmark(?:Start|End)\b/;

function isBookmarkMarker(node: PMNode): boolean {
  if (node.type.name === "bookmarkBlock") return true;
  return (
    node.type.name === "rawInline" &&
    typeof node.attrs.xml === "string" &&
    BOOKMARK_XML.test(node.attrs.xml)
  );
}

function bookmarkSignature(node: PMNode): string {
  return node.type.name === "bookmarkBlock"
    ? `block:${node.attrs.srcId}:${node.attrs.name}`
    : `inline:${node.attrs.xml}`;
}

export const bookmarkGuard = preservedNodeGuard(
  "bookmark",
  isBookmarkMarker,
  bookmarkSignature
);

function isNoteReference(node: PMNode): boolean {
  return node.type.name === "noteReference";
}

export const noteGuard = preservedNodeGuard("note", isNoteReference, (node) =>
  JSON.stringify(node.attrs)
);
