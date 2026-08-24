/**
 * The line spacing of a paragraph.
 *
 * It is set to one of the presets the toolbar offers.
 * All XML handling belongs to `docx/paraProps`; this file only decides which paragraphs to
 * change and to what.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { withLineSpacing } from "../../docx/paraProps";
import {
  type DocumentDefaults,
  type LineSpacing,
  toParagraphFormat,
} from "../../model/format";
import { documentDefaults } from "../documentStyles";
import {
  editableParagraphs,
  editParagraphs,
  paragraphPPr,
  selectedParagraphs,
} from "../paragraphEdits";

/** One line of text, which is what a paragraph nobody gave a line spacing is drawn with */
export const SINGLE_LINE_SPACING: LineSpacing = { rule: "auto", lines: 1 };

/**
 * The line spacing this paragraph is drawn with.
 * What it wrote down, and failing that what its style gives; failing that the document
 * default, and failing that a single line.
 */
function lineSpacingOf(node: PMNode, defaults: DocumentDefaults): LineSpacing {
  return (
    toParagraphFormat(node.attrs.format)?.lineSpacing ??
    defaults.lineSpacing ??
    SINGLE_LINE_SPACING
  );
}

function sameLineSpacing(a: LineSpacing, b: LineSpacing): boolean {
  if (a.rule === "auto") return b.rule === "auto" && a.lines === b.lines;
  return b.rule === a.rule && a.pt === b.pt;
}

/**
 * Sets the line spacing of the selected paragraphs.
 * A paragraph already drawn with that spacing is left untouched, so its original XML survives.
 */
export function setLineSpacing(spacing: LineSpacing): Command {
  return (state, dispatch) => {
    const defaults = documentDefaults(state);
    return editParagraphs(state, dispatch, (node, styles, defaultStyleId) =>
      sameLineSpacing(lineSpacingOf(node, defaults), spacing)
        ? null
        : withLineSpacing(paragraphPPr(node), spacing, styles, defaultStyleId)
    );
  };
}

/** The line spacing the selected paragraphs share. Null when several spacings are mixed or nothing is selected */
export function activeLineSpacing(state: EditorState): LineSpacing | null {
  const defaults = documentDefaults(state);
  const spacings = selectedParagraphs(state).map(({ node }) =>
    lineSpacingOf(node, defaults)
  );
  const first = spacings[0];
  if (!first) return null;
  return spacings.every((spacing) => sameLineSpacing(spacing, first))
    ? first
    : null;
}

/**
 * Whether the selection holds a paragraph to space out at all. The toolbar disables the menu when it
 * holds none, a selection the lock leaves nothing of included (`editor/paragraphEdits`)
 */
export function canSetLineSpacing(state: EditorState): boolean {
  return editableParagraphs(state).length > 0;
}
