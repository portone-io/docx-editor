/**
 * The guards over what a document is opened with and the editor only preserves: the fragments a
 * file falls apart without, the reference standing where a footnote or an endnote is called, and
 * the paragraph that ends a section.
 *
 * The first of those is the two ends of a bookmark range, the two ends of a permission or move
 * range, and the pieces of a field, which say what they say only in the order they stand in. None
 * of it is content the editor writes, so nothing in it can put such a thing back once it is gone,
 * and a document that lost one of a bookmark's two ends cannot be written back as a file at all
 * (`docx/exportDocx` refuses it). What the first two guards hold is therefore the whole list of
 * the nodes they answer for, in the order the document carries them; a section break, which an
 * edit may legitimately move from one paragraph to another, is held by its number instead.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  AddNodeMarkStep,
  AttrStep,
  RemoveMarkStep,
  RemoveNodeMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
} from "prosemirror-transform";
import { W_NS } from "../ooxml/names";
import { parseProps, propsChild } from "../ooxml/props";
import {
  COMMENT_RANGE_MARKERS,
  PERMISSION_MARKERS,
  RANGE_MARKERS,
} from "../ooxml/rangeMarkers";
import { parseAttrs } from "../ooxml/tagScan";
import {
  type ChangeGuard,
  type EditGuardName,
  rangeHolds,
  transactionReaches,
} from "./editGuard";
import { visitPreservedFragments } from "./preservedFragments";

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

/**
 * Whether the deletion guard answers for this fragment.
 *
 * `docx/importPolicy` decides that when the document is opened and bakes the answer into the
 * node, so the rule lives in one place and this layer reads it rather than matching element names
 * against a pattern of its own.
 */
function isGuardedFragment(node: PMNode): boolean {
  return node.attrs.guarded === true;
}

/**
 * A block opened from the body carries no XML of its own but names the fragment it stands for
 * (`docx/importPreserved`), so what tells two markers apart is read from both.
 */
function preservedSignature(node: PMNode): string {
  return `${node.type.name}:${node.attrs.name}:${node.attrs.srcId}:${node.attrs.xml}`;
}

const guardedNodes = preservedNodeGuard(
  "preserved",
  isGuardedFragment,
  preservedSignature
);

const rangeMarkers = new Set([
  ...RANGE_MARKERS,
  ...PERMISSION_MARKERS,
  ...COMMENT_RANGE_MARKERS,
]);

/** Table attributes hold both range markers and disposable producer traces. */
function markerFragments(xml: string): string[] {
  const children = parseProps(`<markers>${xml}</markers>`)?.children ?? [];
  return children.flatMap((child) => {
    if (!rangeMarkers.has(child.name)) return [];
    const props = parseProps(child.xml);
    if (!props) return [];
    const prefixEnd = props.tag.indexOf(":");
    const binding =
      prefixEnd < 0 ? "xmlns" : `xmlns:${props.tag.slice(0, prefixEnd)}`;
    const namespace = parseAttrs(props.attrs ?? "")?.find(
      ([name]) => name === binding
    )?.[1];
    // An inherited binding is supplied by the original document. Explicit foreign bindings
    // do not turn a foreign element with the same local name into a Word range marker.
    if (namespace !== undefined && namespace !== W_NS) return [];
    return [child.xml];
  });
}

function hasMarkerAttrs(node: PMNode): boolean {
  return (
    ((node.type.name === "table" || node.type.name === "tableRow") &&
      typeof node.attrs.leadingXml === "string") ||
    ((node.type.name === "tableCell" || node.type.name === "tableRow") &&
      typeof node.attrs.trailingXml === "string")
  );
}

/** A text edit inside a carrier leaves its XML attributes standing. */
function carrierBoundaryIn(doc: PMNode, from: number, to: number): boolean {
  let found = false;
  doc.nodesBetween(from, to, (node, pos) => {
    if (hasMarkerAttrs(node) && (from <= pos || to >= pos + node.nodeSize)) {
      found = true;
    }
    return !found;
  });
  return found;
}

function markerAttrsReached(tr: Transaction): boolean {
  return tr.steps.some((step, index) => {
    const before = tr.docs[index];
    const after = tr.docs[index + 1] ?? tr.doc;
    if (
      step instanceof AttrStep ||
      step instanceof AddNodeMarkStep ||
      step instanceof RemoveNodeMarkStep
    ) {
      const old = before.nodeAt(step.pos);
      const next = after.nodeAt(step.pos);
      return (
        (old !== null && hasMarkerAttrs(old)) ||
        (next !== null && hasMarkerAttrs(next))
      );
    }
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep)
      return false;
    if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep))
      return true;
    let reached = false;
    step.getMap().forEach((oldStart, oldEnd, newStart, newEnd) => {
      reached ||=
        carrierBoundaryIn(before, oldStart, oldEnd) ||
        carrierBoundaryIn(after, newStart, newEnd);
    });
    return reached;
  });
}

function preservedFragments(doc: PMNode): string[] {
  const found: string[] = [];
  visitPreservedFragments(doc, (node, _pos, xml, attr) => {
    if (attr !== undefined && xml !== null) {
      found.push(
        ...markerFragments(xml).map((fragment) => `marker:${fragment}`)
      );
    } else if (isGuardedFragment(node)) {
      found.push(preservedSignature(node));
    }
  });
  return found;
}

export const preservedGuard: ChangeGuard = {
  name: "preserved",
  change: (tr) =>
    !(transactionReaches(tr, isGuardedFragment) || markerAttrsReached(tr)) ||
    same(preservedFragments(tr.before), preservedFragments(tr.doc)),
  shuts: (intent, state) => {
    if (guardedNodes.shuts(intent, state)) return true;
    if (intent.kind !== "replace") return false;
    let shut = false;
    state.doc.nodesBetween(intent.from, intent.to, (node, pos) => {
      if (intent.from > pos && intent.to < pos + node.nodeSize) return true;
      if (!hasMarkerAttrs(node)) return true;
      for (const attr of ["leadingXml", "trailingXml"]) {
        const xml = node.attrs[attr];
        if (typeof xml === "string" && markerFragments(xml).length > 0)
          shut = true;
      }
      return !shut;
    });
    return shut;
  },
};

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
  shuts: (intent, state) => {
    if (intent.kind !== "replace") return false;
    const $from = state.doc.resolve(intent.from);
    const $to = state.doc.resolve(intent.to);
    // Replacing inline content leaves the paragraph and its section properties standing.
    if ($from.parent.type.name === "paragraph" && $from.sameParent($to)) {
      return false;
    }
    return rangeHolds(state.doc, intent.from, intent.to, endsASection);
  },
};
