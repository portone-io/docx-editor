/**
 * How a block the editor never modelled goes back out, wherever in the document it stands.
 *
 * A block opened under the body names the fragment it stands for and the session holds the bytes;
 * one opened inside a table cell was never a fragment of its own and carries its XML along. An
 * edit may move either into the other's place, so the body writer and the cell writer read both
 * the same way, from here.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import type { ExportRefs } from "./exportRefs";
import { originalBlock, type SessionStore, splitBlockKey } from "./session";

/** Why there is no original to write, which a block that came in from another document answers differently */
export function lostOriginal(
  node: PMNode,
  session: SessionStore | null
): string {
  const srcId: unknown = node.attrs.srcId;
  const key = typeof srcId === "string" ? splitBlockKey(srcId) : null;
  return key === null || key.sessionId === session?.sessionId
    ? "a preserved block has lost its original XML"
    : `a preserved block comes from another document (${key.sessionId})`;
}

export function serializePreservedBlock(
  node: PMNode,
  refs: ExportRefs
): string {
  const xml: unknown = node.attrs.xml;
  if (typeof xml === "string") return xml;
  const imported = refs.session ? originalBlock(node, refs.session) : undefined;
  if (!imported) {
    throw new DocxExportError(
      "lost-original",
      lostOriginal(node, refs.session)
    );
  }
  return imported.xml;
}
