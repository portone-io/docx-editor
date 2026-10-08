import type { Node as PMNode } from "prosemirror-model";

/** Whether this is a run holding no characters, which the file left as a place to write into */
export function isEmptyRun(node: PMNode | null | undefined): boolean {
  return node?.type.name === "emptyRun";
}
