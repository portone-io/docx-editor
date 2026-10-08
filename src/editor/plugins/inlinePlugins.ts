import type { Plugin } from "prosemirror-state";
import { emptyRunPointer } from "./emptyRunPointer";
import { tabCaret } from "./tabCaret";
import { tabDecorations } from "./tabDecorations";
import { tabLayout } from "./tabLayout";
import { tabPointer } from "./tabPointer";

/**
 * The plugins that draw the inline leaves of a story and place the caret around them, shared by the
 * body and every side story so a leaf behaves alike wherever it stands.
 */
export function inlinePlugins(): Plugin[] {
  return [
    // Adjacent text tabs still need separate DOM ranges for layout and pointer selection
    tabDecorations(),
    tabPointer(),
    tabLayout(),
    tabCaret(),
    emptyRunPointer(),
  ];
}
