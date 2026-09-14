/**
 * A side story drawn as markup no editor view stands behind: a note the caret is not in, and later
 * a comment body or a header preview.
 *
 * The runs are drawn by the same spec the editor's run view draws (`editor/views/runMarkView`), so
 * a story set down here wraps where it would in an editor, and what one kind of story draws
 * differently comes in as a node spec. Nothing here knows which kind of story it is drawing, or
 * where the markup goes.
 *
 * The schema writes the document's own source onto what it draws (`w:pPr`, a run's `w:rPr`, a
 * preserved fragment's XML) for an editor to read back, and nothing reads this markup back. So
 * none of it is kept: a selection the browser copies out of a note, with no copy serializer of the
 * editor behind it, then carries no more of the document than a copy out of the editor does
 * (`schema/clipboard`).
 */

import {
  type DOMOutputSpec,
  DOMSerializer,
  type Mark,
  type Node as PMNode,
} from "prosemirror-model";
import { runMarkSpec } from "../../schema";
import type { FontFallbacks } from "../../styles/fontStack";

/** How a node type is drawn in place of the way the schema draws it, by node type name */
export type StoryNodeSpecs = Readonly<
  Record<string, (node: PMNode) => DOMOutputSpec>
>;

export interface StoryMarkupOptions {
  readonly fontFallbacks: FontFallbacks;
  readonly nodeSpecs?: StoryNodeSpecs;
}

/**
 * The data attributes `styles/editor.css` reads on story content, which hold no source: how a
 * preserved fragment is displayed (`data-display`), and whether a run inside a link underlines
 * itself (`data-underline`)
 */
const STYLED_ATTRIBUTES: ReadonlySet<string> = new Set([
  "data-display",
  "data-underline",
]);

function withoutSource(markup: DocumentFragment): DocumentFragment {
  markup.querySelectorAll("*").forEach((element) => {
    for (const name of element.getAttributeNames()) {
      if (name.startsWith("data-") && !STYLED_ATTRIBUTES.has(name)) {
        element.removeAttribute(name);
      }
    }
  });
  return markup;
}

export function storyMarkup(
  story: PMNode,
  { fontFallbacks, nodeSpecs = {} }: StoryMarkupOptions
): DocumentFragment {
  const schema = story.type.schema;
  const serializer = new DOMSerializer(
    { ...DOMSerializer.nodesFromSchema(schema), ...nodeSpecs },
    {
      ...DOMSerializer.marksFromSchema(schema),
      run: (mark: Mark) => runMarkSpec(mark.attrs, fontFallbacks),
    }
  );
  const markup = document.createDocumentFragment();
  serializer.serializeFragment(story.content, {}, markup);
  return withoutSource(markup);
}
