import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { preservedXml, serializeParagraph } from "./serializeParagraph";
import { serializeTable } from "./serializeTable";
import { originalBlock, type SessionStore } from "./session";

/** For a block whose content we never modelled, there is no way to write it back other than the original fragment */
function serializeRaw(node: PMNode, session: SessionStore): string {
  const imported = originalBlock(node, session);
  if (!imported) {
    throw new DocxExportError(
      "lost-original",
      "a preserved block has lost its original XML"
    );
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
  if (node.type.name === "rawBlock") return preservedXml(node);
  if (node.type.name === "bookmarkBlock") return serializeRaw(node, session);
  if (node.type.name === "docxRaw") return serializeRaw(node, session);
  throw new DocxExportError(
    "unsupported-content",
    `block we cannot serialize: ${node.type.name}`
  );
}
