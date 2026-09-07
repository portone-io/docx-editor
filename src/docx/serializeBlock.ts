import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { preservedXml, serializeParagraph } from "./serializeParagraph";
import { serializeTable } from "./serializeTable";
import { originalBlock, type SessionStore, splitBlockKey } from "./session";

/** Why there is no original to write, which a block that came in from another document answers differently */
export function lostOriginal(node: PMNode, session: SessionStore): string {
  const srcId: unknown = node.attrs.srcId;
  const key = typeof srcId === "string" ? splitBlockKey(srcId) : null;
  return key === null || key.sessionId === session.sessionId
    ? "a preserved block has lost its original XML"
    : `a preserved block comes from another document (${key.sessionId})`;
}

/** For a block whose content we never modelled, there is no way to write it back other than the original fragment */
function serializeRaw(node: PMNode, session: SessionStore): string {
  const imported = originalBlock(node, session);
  if (!imported) {
    throw new DocxExportError("lost-original", lostOriginal(node, session));
  }
  return imported.xml;
}

export function serializeBlock(
  node: PMNode,
  session: SessionStore,
  refs: ExportRefs = NO_EXPORT_REFS
): string {
  if (node.type.name === "paragraph") return serializeParagraph(node, refs);
  if (node.type.name === "table") return serializeTable(node, refs);
  // The one preserved block that carries its XML on itself rather than pointing at the original
  if (node.type.name === "rawBlock") return preservedXml(node);
  if (node.type.isInGroup("preserved")) return serializeRaw(node, session);
  throw new DocxExportError(
    "unsupported-content",
    `block we cannot serialize: ${node.type.name}`
  );
}
