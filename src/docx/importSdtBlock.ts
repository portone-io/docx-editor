/**
 * Reading a block-level content control (`w:sdt` under the body, a cell, or another control) as a
 * container of the blocks it holds.
 *
 * The blocks inside a control are the ones the level around it takes, so the caller hands over the
 * reader for its own level rather than this module reaching for one: a control may hold a table
 * while a table may hold a control, and neither may import the other.
 *
 * null leaves the control to whatever held it, which preserves it whole as it was before.
 */

import type { Node as PMNode } from "prosemirror-model";
import { elementChildren } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { controlAttrs, OWN_CONTROL_ATTRS } from "../schema/controlAttrs";
import { controlFactsFrom, modelsBlockContent, readSdtWrapper } from "./sdt";
import { nextKey } from "./wrappers";

/** How the level around a control reads one block of it */
export type BlockReader = (el: Element) => PMNode;

/**
 * One `w:sdt` as a container node, or null for a control this editor keeps whole.
 *
 * `srcId` names the fragment of the session the control was sliced as; a control inside a cell or
 * inside another control was never a fragment of its own and is handed null, as a nested table is.
 *
 * A control whose `w:sdtContent` is missing or empty is kept whole too: `block+` has no way to say
 * "nothing", and creating a paragraph for it would write a blank line the file never held.
 */
export function buildSdtBlock(
  el: Element,
  srcId: string | null,
  readBlock: BlockReader
): PMNode | null {
  if (!modelsBlockContent(el)) return null;
  const wrapper = readSdtWrapper(el);
  if (!wrapper) return null;

  const children = elementChildren(wrapper.content);
  if (children.length === 0) return null;

  return docxSchema.nodes.sdtBlock.create(
    {
      srcId,
      key: nextKey(docxSchema.marks.sdt.name, el),
      ...controlAttrs(OWN_CONTROL_ATTRS, controlFactsFrom(wrapper)),
    },
    children.map(readBlock)
  );
}
