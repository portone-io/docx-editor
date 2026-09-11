/**
 * What a footnote or an endnote draws differently from any other story.
 *
 * A note's story opens with a mark of its own (`w:footnoteRef`, `w:endnoteRef`), which the file
 * keeps as a preserved chip. Word draws that mark as the number the reference carries, so it is
 * drawn here as that label in superscript, the way the reference in the body is.
 */

import type { DOMOutputSpec, Node as PMNode } from "prosemirror-model";
import { docxSchema } from "../../schema";
import { editorClassNames } from "../../styles/classNames";
import type { StoryNodeSpecs } from "../stories/storyMarkup";

const OWN_REFERENCE_MARKS: readonly unknown[] = ["footnoteRef", "endnoteRef"];

function drawnBySchema(node: PMNode): DOMOutputSpec {
  return docxSchema.nodes.rawRunContent.spec.toDOM?.(node) ?? ["span"];
}

/** The node specs a note's story is drawn with, its own reference mark drawn as `labelOf` answers */
export function noteNodeSpecs(labelOf: () => string): StoryNodeSpecs {
  return {
    rawRunContent: (node) =>
      OWN_REFERENCE_MARKS.includes(node.attrs.element)
        ? ["sup", { class: editorClassNames.noteMark }, labelOf()]
        : drawnBySchema(node),
  };
}
