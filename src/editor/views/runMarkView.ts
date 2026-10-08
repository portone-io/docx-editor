/**
 * Draws the run spans inside one editor.
 *
 * The fallback fonts belong to the editor, not to the schema: the schema is one
 * shared value, so its `toDOM` can only ever draw with the default fallbacks.
 * A mark view is where the view gets to draw a run itself, so the fallbacks handed
 * to this editor are the ones that apply.
 *
 * The span is built from the same spec `toDOM` returns and rendered by the same
 * renderer, so with the default fallbacks the DOM is what the schema would have
 * drawn, but for one attribute: a run whose own properties paint a background says
 * so, which draws a run holding no characters inside it as a blank
 * (`styles/editor.css`). Its own properties are read apart from what it inherits by
 * the import's reader, which the schema stands below.
 */

import { DOMSerializer } from "prosemirror-model";
import type { MarkViewConstructor } from "prosemirror-view";
import { ownBackground } from "../../docx/runProps";
import { runMarkSpec } from "../../schema";
import { editorAttributes } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";

export function runMarkView(fontFallbacks: FontFallbacks): MarkViewConstructor {
  return (mark) => {
    const view = DOMSerializer.renderSpec(
      document,
      runMarkSpec(mark.attrs, fontFallbacks)
    );
    if (ownBackground(mark) !== null && view.dom instanceof HTMLElement) {
      view.dom.setAttribute(editorAttributes.paintsBackground, "");
    }
    return view;
  };
}
