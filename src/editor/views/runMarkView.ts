/**
 * Draws the run spans inside one editor.
 *
 * The fallback fonts belong to the editor, not to the schema: the schema is one
 * shared value, so its `toDOM` can only ever draw with the default fallbacks.
 * A mark view is where the view gets to draw a run itself, so the fallbacks handed
 * to this editor are the ones that apply.
 *
 * The span is built from the same spec `toDOM` returns and rendered by the same
 * renderer, so with the default fallbacks the DOM is exactly what the schema would
 * have drawn.
 */

import { DOMSerializer } from "prosemirror-model";
import type { MarkViewConstructor } from "prosemirror-view";
import { runMarkSpec } from "../../schema";
import type { FontFallbacks } from "../../styles/fontStack";

export function runMarkView(fontFallbacks: FontFallbacks): MarkViewConstructor {
  return (mark) =>
    DOMSerializer.renderSpec(document, runMarkSpec(mark.attrs, fontFallbacks));
}
