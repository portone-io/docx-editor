/**
 * Walking the blocks of a story the way the sequence the document writes them in runs.
 *
 * `CT_SdtContentBlock` (§17.5.2.34) holds what `CT_Body` holds, so a paragraph inside a
 * block-level content control is one of the story's own paragraphs and is reached here, at any
 * depth of nesting. A table cell's paragraphs are not that sequence - a table is one block of the
 * story and what stands in its cells belongs to the cells - so a table is not walked into.
 *
 * Everything that has to name the last, the first or every paragraph of a story asks here, so the
 * readers cannot come to disagree about what a control is.
 */

import type { Node as PMNode } from "prosemirror-model";
import { docxSchema } from "../schema";
import { isBlockControl } from "../schema/controlAttrs";

function isParagraph(node: PMNode): boolean {
  return node.type === docxSchema.nodes.paragraph;
}

/**
 * Every paragraph one block of a story stands for, in document order: the block itself where it is
 * a paragraph, and otherwise the paragraphs of the content control around them.
 *
 * `pos` is where the block itself stands, and each paragraph is visited with its own position.
 */
export function eachStoryParagraph(
  block: PMNode,
  pos: number,
  visit: (paragraph: PMNode, paragraphPos: number) => void
): void {
  if (isParagraph(block)) {
    visit(block, pos);
    return;
  }
  if (!isBlockControl(block)) return;
  block.forEach((child, offset) => {
    eachStoryParagraph(child, pos + 1 + offset, visit);
  });
}

/** The paragraphs a whole story draws, in order */
export function storyParagraphs(story: PMNode): readonly PMNode[] {
  const found: PMNode[] = [];
  story.forEach((block, offset) => {
    eachStoryParagraph(block, offset, (paragraph) => found.push(paragraph));
  });
  return found;
}

/**
 * The one paragraph a story holds and nothing besides, and null for every other story. A control
 * around it is not content of its own, so a story holding one is as empty as its paragraph is.
 */
export function onlyStoryParagraph(container: PMNode): PMNode | null {
  const only = container.childCount === 1 ? container.firstChild : null;
  if (only === null) return null;
  if (isParagraph(only)) return only;
  return isBlockControl(only) ? onlyStoryParagraph(only) : null;
}

/** The path to the last paragraph of a story, through the controls that hold it, or null for none */
function lastParagraphPath(container: PMNode): readonly number[] | null {
  for (let index = container.childCount - 1; index >= 0; index -= 1) {
    const child = container.child(index);
    if (isParagraph(child)) return [index];
    if (!isBlockControl(child)) continue;
    const deeper = lastParagraphPath(child);
    if (deeper !== null) return [index, ...deeper];
  }
  return null;
}

function replaceAtPath(
  container: PMNode,
  path: readonly number[],
  written: PMNode
): PMNode {
  const [index, ...rest] = path;
  if (index === undefined) return written;
  return container.copy(
    container.content.replaceChild(
      index,
      rest.length === 0
        ? written
        : replaceAtPath(container.child(index), rest, written)
    )
  );
}

/**
 * The story with its last paragraph rewritten, wherever the controls around it put it. The story
 * itself comes back where it holds no paragraph, or where `rewrite` answers null for the one it
 * found.
 */
export function withLastStoryParagraph(
  story: PMNode,
  rewrite: (paragraph: PMNode) => PMNode | null
): PMNode {
  const path = lastParagraphPath(story);
  if (path === null) return story;
  const paragraph = path.reduce<PMNode>(
    (node, index) => node.child(index),
    story
  );
  const written = rewrite(paragraph);
  return written === null ? story : replaceAtPath(story, path, written);
}
