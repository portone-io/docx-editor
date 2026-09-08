/**
 * The lists that were started during editing, and the part their definitions are written into.
 *
 * Both the writer and the export invariants ask these questions, so they stand apart from the
 * writer rather than being read off it.
 */

import type { Node as PMNode } from "prosemirror-model";
import { toParagraphFormat } from "../model/format";
import { NEW_LISTS_ATTR, newListsOf } from "../numbering/listRegistry";
import { type NewList, parseNumbering } from "../numbering/parseNumbering";
import { R_NS } from "../ooxml/xml";
import { CONTENT_TYPES_PATH } from "./packageParts";
import type { SessionStore } from "./session";

/**
 * How the main part relates the numbering part, and what the package declares that part to hold.
 *
 * The reader finds the part by this relationship and the writer relates a new one under it, so
 * neither can name it differently from the other.
 */
export const NUMBERING_REL_TYPE = `${R_NS}/numbering`;

export const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

/** Collects the numbering ids used by this block and by the paragraphs inside it (down into table cells) */
function collectNumIds(node: PMNode, into: Set<number>): void {
  const visit = (candidate: PMNode): boolean => {
    if (candidate.type.name !== "paragraph") return true;
    const numId = toParagraphFormat(candidate.attrs.format)?.numbering?.numId;
    if (numId !== undefined) into.add(numId);
    return false;
  };
  if (visit(node)) node.descendants(visit);
}

function numIdsIn(node: PMNode): Set<number> {
  const used = new Set<number>();
  collectNumIds(node, used);
  return used;
}

function numIdsAtOpen(session: SessionStore): Set<number> {
  const used = new Set<number>();
  for (const block of session.blocks) collectNumIds(block.node, used);
  return used;
}

/**
 * The numbering ids that appeared during editing and have no definition.
 *
 * An id that was already in use without a definition when the document was opened is left alone,
 * because exporting such a document without editing it must not disturb numbering.xml.
 */
export function newNumIds(doc: PMNode, session: SessionStore): number[] {
  const defined = parseNumbering(session.numberingXml).lists;
  const atOpen = numIdsAtOpen(session);
  return Array.from(numIdsIn(doc))
    .filter((numId) => !defined.has(numId) && !atOpen.has(numId))
    .sort((a, b) => a - b);
}

/** The lists started during editing, split by whether a definition was registered for each */
export interface StartedLists {
  /** The definition each list was started with, which is what goes into numbering.xml */
  readonly defined: ReadonlyMap<number, NewList>;
  /** The numbers a list was started under that no definition stands behind */
  readonly unregistered: readonly number[];
}

/**
 * The lists started during editing, each against the definition it was started with.
 *
 * The definition is the one the document node carries (`numbering/listRegistry`), which the list
 * commands and the clipboard record as they start a list. A number carrying none is reported here
 * rather than written bare, which the export invariants turn into a refusal.
 */
export function startedLists(doc: PMNode, session: SessionStore): StartedLists {
  const registered = newListsOf(doc.attrs[NEW_LISTS_ATTR]);
  const defined = new Map<number, NewList>();
  const unregistered: number[] = [];
  for (const numId of newNumIds(doc, session)) {
    const list = registered.get(numId);
    if (list === undefined) unregistered.push(numId);
    else defined.set(numId, list);
  }
  return { defined, unregistered };
}

/** The numbering part as it was opened, which a new definition is spliced into. null when the document has none */
export function numberingPartOf(
  session: SessionStore
): { path: string; bytes: Uint8Array } | null {
  const path = session.numberingPartPath;
  const bytes = path === null ? undefined : session.parts.get(path);
  return path === null || bytes === undefined ? null : { path, bytes };
}

/**
 * Whether the definition of a new list has somewhere to go: the numbering part the document was
 * opened with, or one written from scratch, which only a package that can declare what the new
 * part holds may take.
 */
export function canDefineNewList(session: SessionStore): boolean {
  return (
    numberingPartOf(session) !== null || session.parts.has(CONTENT_TYPES_PATH)
  );
}
