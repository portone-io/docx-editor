/**
 * The lists that were started during editing, and the part their definitions are written into.
 *
 * Both the writer and the export invariants ask these questions, so they stand apart from the
 * writer rather than being read off it.
 */

import type { Node as PMNode } from "prosemirror-model";
import { toParagraphFormat } from "../model/format";
import { parseNumbering } from "../numbering/parseNumbering";
import type { SessionStore } from "./session";

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

/** The numbering part as it was opened, which a new definition is spliced into. null when the document has none */
export function numberingPartOf(
  session: SessionStore
): { path: string; bytes: Uint8Array } | null {
  const path = session.numberingPartPath;
  const bytes = path === null ? undefined : session.parts.get(path);
  return path === null || bytes === undefined ? null : { path, bytes };
}
