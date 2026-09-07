/**
 * The guards over what a document is opened with and the editor only preserves: the two ends of a
 * bookmark range, the reference standing where a footnote or an endnote is called, and the
 * paragraph that ends a section.
 *
 * None of it is content the editor writes, so nothing in it can put such a thing back once it is
 * gone, and a document that lost one of a bookmark's two ends cannot be written back as a file at
 * all (`docx/exportDocx` refuses it). What the first two guards hold is therefore the whole list of
 * the nodes they answer for, in the order the document carries them; a section break, which an
 * edit may legitimately move from one paragraph to another, is held by its number instead.
 */

import type { Node as PMNode } from "prosemirror-model";
import { parseProps, propsChild } from "../ooxml/props";
import {
  type ChangeGuard,
  type EditGuardName,
  rangeHolds,
  transactionReaches,
} from "./editGuard";

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
  name: EditGuardName,
  holds: (node: PMNode) => boolean,
  signature: Signature
): ChangeGuard {
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

/**
 * Whether this paragraph's own properties end a section.
 *
 * `parseProps` counts depth and lists direct children alone, so the `w:sectPr` a `w:pPrChange`
 * carries - the properties the paragraph wore before a tracked change - is not a break this
 * paragraph lays down.
 */
function endsASection(node: PMNode): boolean {
  if (node.type.name !== "paragraph") return false;
  const pPr = node.attrs.pPr;
  if (typeof pPr !== "string") return false;
  const props = parseProps(pPr);
  return props !== null && propsChild(props.children, "sectPr") !== undefined;
}

/** How many paragraphs of this document lay down a section break */
function sectionBreaks(doc: PMNode): number {
  let seen = 0;
  doc.descendants((node) => {
    if (endsASection(node)) seen += 1;
    return true;
  });
  return seen;
}

/**
 * A guard that lets a change through only while it leaves at least as many section breaks standing
 * as it found.
 *
 * A paragraph-level `w:sectPr` is not content: it is where one section of the document ends, and
 * the section carries the page size, the margins and the headers everything up to it is laid out
 * under. Joining that paragraph into the one above it - a Backspace at its start, a selection run
 * across its boundary - would take the whole section away with it, and nothing the editor writes
 * puts a section back.
 *
 * It counts rather than compares, so splitting such a paragraph is still allowed: Enter leaves the
 * break on the half that ends up last (`docx/cloning`), and one break stands where one stood. An
 * edit that adds a break is no business of this guard's either.
 *
 * The count walks the whole document, so it is reached for only once a step has touched a
 * paragraph that lays a break down.
 */
export const sectionGuard: ChangeGuard = {
  name: "section",
  change: (tr) =>
    !transactionReaches(tr, endsASection) ||
    sectionBreaks(tr.doc) >= sectionBreaks(tr.before),
  shuts: (intent, state) =>
    intent.kind === "replace" &&
    rangeHolds(state.doc, intent.from, intent.to, endsASection),
};
