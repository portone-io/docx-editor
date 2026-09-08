import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { sameSource } from "../schema/sourceEquality";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { serializeParagraph } from "./serializeParagraph";
import { serializePreservedBlock } from "./serializePreserved";
import { serializeTable } from "./serializeTable";
import { originalBlock } from "./session";

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

/**
 * An unchanged block is written back as its original XML; only a changed block is rebuilt.
 *
 * Unchanged is judged by `sameSource` rather than by `Node.eq`, because opening a file works the
 * display attrs out again (`schema/attrRoles`) and a block rebuilt over that would lose the markup
 * the writer does not model, the properties of a cell continuing a vertical merge among it.
 *
 * `onlyBlock` says the block stands alone in its story. An original of empty XML belongs to the
 * editable placeholder a story with no block of its own opens with; once another block joins it,
 * that paragraph stands for a real blank line and has to be written.
 */
export function blockXml(
  node: PMNode,
  refs: ExportRefs,
  onlyBlock: boolean
): string {
  const imported =
    refs.session === null ? undefined : originalBlock(node, refs.session);
  if (
    imported &&
    (imported.xml !== "" || onlyBlock) &&
    sameSource(node, imported.node)
  ) {
    return imported.xml;
  }
  return serializeBlock(node, refs);
}
