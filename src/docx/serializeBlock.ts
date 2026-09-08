import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { serializeParagraph } from "./serializeParagraph";
import { serializePreservedBlock } from "./serializePreserved";
import { serializeTable } from "./serializeTable";

export function serializeBlock(
  node: PMNode,
  refs: ExportRefs = NO_EXPORT_REFS
): string {
  if (node.type.name === "paragraph") return serializeParagraph(node, refs);
  if (node.type.name === "table") return serializeTable(node, refs);
  if (node.type.isInGroup("preserved")) {
    return serializePreservedBlock(node, refs);
  }
  throw new DocxExportError(
    "unsupported-content",
    `block we cannot serialize: ${node.type.name}`
  );
}
