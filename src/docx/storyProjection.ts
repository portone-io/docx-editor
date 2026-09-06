/**
 * The document story as the parts of the package see it: which blocks the editor models and
 * writes back itself, and the text those blocks go back out as.
 *
 * Two files saying the same thing can spell it differently. A block carries the XML it arrived
 * as on its attrs, so a table Word wrote (`w:w` before `w:type` inside a `w:tcW`) and the same
 * table this editor wrote (the other way round) compare unequal attr for attr, and a comparison
 * over the model would call that a body edit. Writing both stories out through the one writer
 * takes the wording out of the comparison and leaves what the story says, because the writer
 * reaches a fixed point on its own output: opening what it wrote and writing it again gives the
 * same text (the test beside this file holds that over every fixture).
 *
 * The price is that the comparison levels whatever the writer levels: two blocks the writer would
 * put out alike read alike here, however differently the two files worded them. Everything the
 * writer carries through is compared, and what falls out of the comparison is only what would have
 * been reworded on its way out anyway. The test beside this file holds the differences known to
 * fall under that, one case each, and holds the other half: what the writer carries is compared.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { NO_EXPORT_REFS } from "./exportRefs";
import { serializeBlock } from "./serializeBlock";
import type { SessionStore } from "./session";

/** A document and the package it was opened from, which is what writing it back out takes */
export interface Story {
  readonly doc: PMNode;
  readonly session: SessionStore;
}

/** Whether this is a block the editor takes apart and writes back itself, rather than one it keeps as it came */
export function isModelledBlock(node: PMNode): boolean {
  return node.type.isInGroup("modelled");
}

/**
 * Every top-level block as the writer puts it out, with the markup that is allowed to have
 * changed taken away first.
 *
 * null where a block cannot be written back at all, which is a story no comparison can be held
 * over: a caller judging two files answers such a story the way it answers a changed one.
 */
export function comparableStory(
  story: Story,
  strip: (node: PMNode) => PMNode
): readonly string[] | null {
  const blocks: string[] = [];
  try {
    story.doc.forEach((block) => {
      blocks.push(serializeBlock(strip(block), story.session, NO_EXPORT_REFS));
    });
  } catch (error) {
    if (error instanceof DocxExportError) return null;
    throw error;
  }
  return blocks;
}
