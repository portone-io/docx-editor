/**
 * The values a paragraph draws with, read out of the document's styles along the same path the
 * import takes (`docx/formatting`), so that a paragraph an edit built, one it re-aligned and one
 * the file arrived with all carry the same values for the same properties.
 */

import type { Node as PMNode } from "prosemirror-model";
import { paragraphAttrsFor } from "../../docx/formatting";
import { paragraphPPr } from "../paragraphEdits";
import type { DocumentDeriver } from "./displayDerivation";

/** Whether the paragraph carries no values at all, which is one the styles were never read into */
function carriesNothing(node: PMNode): boolean {
  return node.attrs.format === null && node.attrs.styleRun === null;
}

/**
 * Whether the paragraph can be left as it stands: it is the paragraph it was, its own properties
 * are as they were, and it carries values. A paragraph a composition held back carries none, so it
 * is read into whatever it maps back to.
 */
function stands(node: PMNode, previous: PMNode | null): boolean {
  return (
    previous !== null &&
    previous.attrs.pPr === node.attrs.pPr &&
    !carriesNothing(node)
  );
}

/** Whether the spot being composed in falls inside this paragraph */
function holdsComposition(
  node: PMNode,
  pos: number,
  composing: number | null
): boolean {
  return (
    composing !== null && composing >= pos && composing <= pos + node.nodeSize
  );
}

export const paragraphDisplay: DocumentDeriver = {
  name: "paragraph",
  nodeTypes: ["paragraph"],
  derive(node, pos, _doc, context, previous) {
    if (
      stands(node, previous) ||
      holdsComposition(node, pos, context.composingAt)
    ) {
      return [];
    }
    return [
      {
        pos,
        attrs: paragraphAttrsFor(
          paragraphPPr(node),
          context.document.formatting
        ),
      },
    ];
  },
};
