import type { BlockKind } from "../blockKinds";
import { paragraphKind } from "./paragraphKind";
import { sdtBlockKind } from "./sdtBlockKind";
import { tableKind } from "./tableKind";

/**
 * The block shapes the editor pages, in the order they are asked. A new kind goes ahead of the
 * kinds it is narrower than, and the last one claims every block none of the others did.
 *
 * A content control is a container of blocks, and pages what it holds by the very list the editor
 * was built with (`page/pageDecorations`), which is handed to it with the block: a kind a consumer
 * registered ahead of these claims a block standing inside a control exactly as it claims the same
 * block standing under the body.
 */
export const DEFAULT_BLOCK_KINDS: readonly BlockKind[] = [
  sdtBlockKind,
  tableKind,
  paragraphKind,
];
