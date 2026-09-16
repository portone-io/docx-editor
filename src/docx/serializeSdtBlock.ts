/**
 * How a block-level content control goes back out: the opening tag it arrived with, and the blocks
 * it holds inside the `w:sdtContent` this editor writes (`./sdt`).
 *
 * A control nobody edited never reaches here at all - `blockXml` hands back the bytes it arrived
 * as, the way it does for a table (`./serializeBlock`).
 *
 * A control holds the same blocks as the level it stands in, and a control may stand inside a
 * table as a table may stand inside a control, so the caller hands over the writer for its own
 * level rather than this module reaching for one.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { ExportRefs } from "./exportRefs";
import { sdtXml } from "./sdt";

/** How the level around a control writes one block of it */
export type BlockWriter = (block: PMNode, refs: ExportRefs) => string;

export function serializeSdtBlock(
  node: PMNode,
  refs: ExportRefs,
  writeBlock: BlockWriter
): string {
  return sdtXml(
    node.attrs.sdtPrefix,
    node.children.map((block) => writeBlock(block, refs)).join("")
  );
}

/**
 * A control holding nothing, which is the same opening around a content tag with nothing in it.
 *
 * A control that arrived writing no content element at all gains one here. The two shapes state
 * the same thing - §17.5.2.34 makes the element a cache of what stood inside - and only a control
 * something rewrote reaches this at all; one nobody touched goes out as the bytes it arrived as.
 */
export function serializeSdtEmpty(node: PMNode): string {
  return sdtXml(node.attrs.sdtPrefix, "");
}
