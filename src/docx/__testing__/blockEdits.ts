/**
 * The one-block edits the round-trip and validation suites make: replacing the first text of a
 * body block, so that the export has to rebuild that block and no other.
 *
 * A paragraph and a table are the two block kinds the writer rebuilds down different paths, and
 * a suite that edits only a paragraph never sends a producer's table through the table writer.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import { withBlocks } from "../../__testing__/docx";
import { docxSchema } from "../../schema";

export type EditedBlock = "paragraph" | "table";

/** The index of the first body block of this kind holding text, the one an editing session reaches first */
export function firstBlockIndex(doc: PMNode, kind: EditedBlock): number {
  let index = -1;
  doc.forEach((block, _offset, at) => {
    if (index !== -1) return;
    if (block.type.name === kind && block.textContent !== "") index = at;
  });
  if (index === -1) throw new Error(`the document has no ${kind} to edit`);
  return index;
}

/** The node with its first text node reading `text` instead, or null when nothing under it is text */
function withFirstTextReplaced(node: PMNode, text: string): PMNode | null {
  if (node.isText) return docxSchema.text(text, node.marks);
  const children: PMNode[] = [];
  let replaced = false;
  node.forEach((child) => {
    const next = replaced ? null : withFirstTextReplaced(child, text);
    if (next !== null) replaced = true;
    children.push(next ?? child);
  });
  return replaced ? node.copy(Fragment.from(children)) : null;
}

/** The document with the block at `index` holding `text` where its first text stood */
export function withEditedBlock(
  doc: PMNode,
  index: number,
  text: string
): PMNode {
  const blocks: PMNode[] = [];
  doc.forEach((block, _offset, at) => {
    if (at !== index) {
      blocks.push(block);
      return;
    }
    const edited = withFirstTextReplaced(block, text);
    if (edited === null) throw new Error(`block ${index} holds no text`);
    blocks.push(edited);
  });
  return withBlocks(doc, blocks);
}

/** The document with the first block of this kind edited */
export function withEditedFirst(
  doc: PMNode,
  kind: EditedBlock,
  text: string
): PMNode {
  return withEditedBlock(doc, firstBlockIndex(doc, kind), text);
}
