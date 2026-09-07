import type { BlockKind } from "../blockKinds";
import { paragraphKind } from "./paragraphKind";
import { tableKind } from "./tableKind";

/**
 * The block shapes the editor pages, in the order they are asked. A new kind goes ahead of the
 * kinds it is narrower than, and the last one claims every block none of the others did.
 */
export const DEFAULT_BLOCK_KINDS: readonly BlockKind[] = [
  tableKind,
  paragraphKind,
];
