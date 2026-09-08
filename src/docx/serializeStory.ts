/**
 * Writes a side story back into the element it stands in.
 *
 * The rule is the body's rule (`./serializeBlock`): a story nobody edited goes back out as the
 * bytes it arrived as, and one that was edited is written block by block, each untouched block
 * still verbatim. That is what keeps a comment's second paragraph and its run formatting through
 * an edit of the first, and what keeps an untouched Comments part byte-identical.
 */

import type { Node as PMNode } from "prosemirror-model";
import { sameSource } from "../schema/sourceEquality";
import type { ExportRefs } from "./exportRefs";
import { blockXml } from "./serializeBlock";
import type { ImportedStory } from "./story";

/** The two ends of the element a story stands in, for a story the package did not arrive holding */
export interface StoryContainer {
  open: string;
  close: string;
}

export function serializeStory(
  current: PMNode,
  imported: ImportedStory | null,
  container: StoryContainer,
  refs: ExportRefs
): string {
  const ends: StoryContainer = imported ?? container;
  if (imported && sameSource(current, imported.doc)) {
    return (
      ends.open +
      imported.blocks.map((block) => block.xml).join("") +
      ends.close
    );
  }
  const pieces: string[] = [];
  current.forEach((block) => {
    pieces.push(blockXml(block, refs, current.childCount === 1));
  });
  return ends.open + pieces.join("") + ends.close;
}
