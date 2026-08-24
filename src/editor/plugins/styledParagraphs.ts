/**
 * Applies inherited style display values to paragraphs created in the editor. The paragraph under
 * active IME composition is deferred to avoid redrawing and canceling the composition.
 */

import type { Attrs, Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  effectiveParagraphFormat,
  effectiveParagraphStyle,
  layerRunFormat,
} from "../../docx/formatting";
import { docxSchema } from "../../schema";
import { documentParagraphFormatting } from "../documentStyles";

/** A paragraph the styles have not been read into, and where it stands */
interface ParagraphSpot {
  pos: number;
  node: PMNode;
}

/** Whether an edit created this paragraph without any resolved display values. */
function needsFormatting(node: PMNode): boolean {
  return node.attrs.format === null && node.attrs.styleRun === null;
}

/**
 * The paragraphs whose display values have not been resolved yet.
 * Paragraph text is not walked into: this runs after every edit that changed the document, typing
 * included, so it may not cost more than a walk over the blocks.
 */
function paragraphSpots(doc: PMNode): ParagraphSpot[] {
  const found: ParagraphSpot[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== docxSchema.nodes.paragraph) return true;
    if (needsFormatting(node)) found.push({ pos, node });
    return false;
  });
  return found;
}

/** The attributes the paragraph is to carry, or null where its style lays down nothing to carry */
function styledAttrs(node: PMNode, state: EditorState): Attrs {
  const context = documentParagraphFormatting(state);
  const style = effectiveParagraphStyle(node.attrs.pPr, context);
  const format = effectiveParagraphFormat(node.attrs.pPr, context);
  const styleRun = style ? layerRunFormat(style.run, null) : null;
  return { ...node.attrs, format, styleRun };
}

/**
 * Where the composition the browser has open stands, which is the one spot left alone.
 * Null while nothing is being composed, and while the editor stands outside a view at all.
 */
function composingAt(
  view: EditorView | null,
  state: EditorState
): number | null {
  return view?.composing === true ? state.selection.from : null;
}

/** Whether the spot being composed in falls inside this paragraph */
function holdsComposition(
  paragraph: ParagraphSpot,
  composing: number | null
): boolean {
  return (
    composing !== null &&
    composing >= paragraph.pos &&
    composing <= paragraph.pos + paragraph.node.nodeSize
  );
}

/**
 * Reads the document's styles into the paragraphs an edit built.
 *
 * No transaction of its own goes to the history: ProseMirror hands an appended transaction to the
 * history as part of the same event, so a single undo takes the edit and these values back together.
 *
 * An appended transaction is handed nothing but states, so the view an open composition is read
 * from is the one the plugin is given when it mounts.
 */
export function styledParagraphs(): Plugin {
  let live: EditorView | null = null;
  return new Plugin({
    view(view) {
      live = view;
      return {
        destroy() {
          live = null;
        },
      };
    },
    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      const composing = composingAt(live, newState);
      const tr = newState.tr;
      for (const paragraph of paragraphSpots(newState.doc)) {
        if (holdsComposition(paragraph, composing)) continue;
        const attrs = styledAttrs(paragraph.node, newState);
        if (!paragraph.node.hasMarkup(paragraph.node.type, attrs)) {
          tr.setNodeMarkup(paragraph.pos, null, attrs);
        }
      }
      return tr.docChanged ? tr : null;
    },
  });
}
