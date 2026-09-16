/**
 * Reading a block-level content control (`w:sdt` under the body, a cell, or another control) as a
 * container of the blocks it holds, or as the node a control holding nothing stands as.
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
import {
  type ControlFacts,
  controlAttrs,
  OWN_CONTROL_ATTRS,
} from "../schema/controlAttrs";
import { controlFactsFrom, modelsBlockContent, readSdtContents } from "./sdt";
import { nextKey } from "./wrappers";

/** How the level around a control reads one block of it */
export type BlockReader = (el: Element) => PMNode;

/** The attributes both shapes carry, the control's own facts among them */
function controlNodeAttrs(
  el: Element,
  srcId: string | null,
  facts: ControlFacts
): Record<string, unknown> {
  return {
    srcId,
    key: nextKey(docxSchema.marks.sdt.name, el),
    ...controlAttrs(OWN_CONTROL_ATTRS, facts),
  };
}

/**
 * One `w:sdt` as a container node, or null for a control this editor keeps whole.
 *
 * `srcId` names the fragment of the session the control was sliced as; a control inside a cell or
 * inside another control was never a fragment of its own and is handed null, as a nested table is.
 *
 * A control holding nothing is read as the atom that draws nothing, in either of the two shapes
 * that state it (`./sdt`).
 */
export function buildSdtBlock(
  el: Element,
  srcId: string | null,
  readBlock: BlockReader
): PMNode | null {
  if (!modelsBlockContent(el)) return null;
  const read = readSdtContents(el);
  if (read === null) return null;
  if (read.kind === "nothing") {
    return docxSchema.nodes.sdtEmpty.create(
      controlNodeAttrs(el, srcId, read.facts)
    );
  }
  return docxSchema.nodes.sdtBlock.create(
    controlNodeAttrs(el, srcId, controlFactsFrom(read.wrapper)),
    elementChildren(read.wrapper.content).map(readBlock)
  );
}
