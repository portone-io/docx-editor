/**
 * What a document holds that this editor could not model, as a list a host application can read.
 *
 * Opening a file quietly hides some markup, stands a placeholder in front of the rest, and a
 * writer approximates a little more on the way back out. None of that shows on the page, so this
 * is where it is said out loud: one note per preserved node, plus whatever a write approximated.
 */

import type { Node as PMNode } from "prosemirror-model";
import { localPart } from "../ooxml/xml";
import { splitBlockKey } from "./session";

export type FidelitySeverity =
  /** Carried through whole, with nothing of it on screen */
  | "hidden"
  /** Carried through whole behind a placeholder that stands in its place and cannot be edited */
  | "placeholder"
  /** Written back as something near what arrived rather than as what arrived */
  | "approximated";

export type FidelityCode =
  | "preserved-run-content"
  | "preserved-inline"
  | "preserved-block"
  | "range-marker"
  | "paragraph-demoted"
  | "table-demoted";

export interface FidelityNote {
  severity: FidelitySeverity;
  code: FidelityCode;
  /** The package path of the part this is about, or null when the caller holds no session */
  part: string | null;
  /** Which block of the body as it was opened this falls in. null for a node made while editing */
  block: number | null;
  /** Where the node stands in the document. null for an approximation with no node behind it */
  pos: number | null;
  /** The original element, for example `w:bookmarkStart` */
  element: string;
}

/** Where a writer records an approximation it had to make while a body was written */
export interface FidelityCollector {
  add(note: FidelityNote): void;
}

/** What a serializer called on its own, outside an export, records its approximations to */
export const NO_FIDELITY_COLLECTOR: FidelityCollector = { add: () => {} };

type Preserved = Pick<FidelityNote, "severity" | "code">;

function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/** The name a preserved fragment opens with, for a node holding the XML rather than the name */
function elementNameOf(xml: string): string | null {
  return /^\s*<\s*([^\s/>]+)/.exec(xml)?.[1] ?? null;
}

function elementOf(node: PMNode): string | null {
  return (
    stringAttr(node.attrs.name) ??
    elementNameOf(stringAttr(node.attrs.xml) ?? "")
  );
}

function demotedBlock(element: string | null): Preserved {
  const name = localNameOf(element);
  if (name === "tbl") return { severity: "placeholder", code: "table-demoted" };
  if (name === "p") {
    return { severity: "placeholder", code: "paragraph-demoted" };
  }
  return { severity: "placeholder", code: "preserved-block" };
}

/**
 * How much of the fragment is on screen.
 *
 * A chip stands in front of what it holds and nothing else draws at all, which is the same
 * division `FidelitySeverity` draws between a placeholder and something hidden.
 */
function severityOf(display: unknown): FidelitySeverity {
  return display === "chip" ? "placeholder" : "hidden";
}

/** What a node says about itself, or null for a node the editor models and draws in full */
function preservedBy(node: PMNode, element: string | null): Preserved | null {
  switch (node.type.name) {
    case "docxRaw":
      return demotedBlock(element);
    case "rawBlock":
      return { severity: "placeholder", code: "preserved-block" };
    case "bookmarkBlock":
      return { severity: "hidden", code: "range-marker" };
    case "rawRunContent":
      return {
        severity: severityOf(node.attrs.display),
        code: "preserved-run-content",
      };
    case "rawInline":
      return preservedInline(node, element);
    default:
      return null;
  }
}

/**
 * What a fragment kept beside the runs of a paragraph reports.
 *
 * A range marker is told from the rest by the guard answering for it: a bookmark, a permission
 * range and a move range are the invisible pairs a document falls apart without, and a `w:proofErr`
 * standing beside them is invisible too but nothing depends on it.
 */
function preservedInline(
  node: PMNode,
  element: string | null
): Preserved | null {
  const name = stringAttr(node.attrs.element) ?? localNameOf(element);
  // A run with no children at all is kept whole (`./importParagraph`). Word draws nothing there
  // either, so the file lost nothing and there is nothing to report
  if (name === "r") return null;
  const severity = severityOf(node.attrs.display);
  return severity === "hidden" && node.attrs.guarded === true
    ? { severity, code: "range-marker" }
    : { severity, code: "preserved-inline" };
}

function localNameOf(element: string | null): string | null {
  return element === null ? null : localPart(element);
}

function blockNumberOf(node: PMNode): number | null {
  const srcId = stringAttr(node.attrs.srcId);
  const key = srcId === null ? null : splitBlockKey(srcId);
  return key?.index ?? null;
}

/**
 * One note per preserved node, in document order.
 *
 * The nodes carry the original element and the block they came from themselves, so the list is a
 * projection of the document rather than something an import has to remember alongside it. A
 * document opened and one edited for an hour are read the same way, and a `pos` is enough to take
 * a reader to the node the note is about.
 */
export function fidelityNotesOf(
  doc: PMNode,
  part: string | null
): FidelityNote[] {
  const notes: FidelityNote[] = [];
  let block: number | null = null;
  doc.descendants((node, pos, parent) => {
    if (parent === doc) block = blockNumberOf(node);
    const element = elementOf(node);
    const preserved = preservedBy(node, element);
    if (preserved) {
      notes.push({ ...preserved, part, block, pos, element: element ?? "?" });
    }
    return true;
  });
  return notes;
}
