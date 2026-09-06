/**
 * The document story as the parts of the package see it: which blocks the editor models and
 * writes back itself, and which ones it only carries.
 */

import type { Node as PMNode } from "prosemirror-model";

/** Whether this is a block the editor takes apart and writes back itself, rather than one it keeps as it came */
export function isModelledBlock(node: PMNode): boolean {
  return node.type.isInGroup("modelled");
}
