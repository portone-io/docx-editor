/**
 * Paragraph formatting commands: the alignment, and the style the paragraph points at.
 *
 * Alignment swaps the `w:jc` of the selected paragraphs and leaves the rest of the paragraph
 * formatting untouched.
 * Applying a style writes nothing but the `w:pStyle`, the same as Word: the values the style
 * lays down are inherited rather than copied, so the direct formatting already written keeps
 * beating them.
 * All XML handling belongs to `docx/paraProps`; this file only calls those functions.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import {
  type FormattingContext,
  type ParagraphAttrs,
  paragraphAttrsOf,
  type ResolvedParagraph,
  resolveParagraph,
  runMarkUnder,
  styleIdOf,
} from "../../docx/formatting";
import {
  type ParagraphProps,
  withParagraphAlign,
  withParagraphStyle,
} from "../../docx/paraProps";
import { type ParagraphAlign, toParagraphFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { lockedMarkOf } from "../../schema/locks";
import { documentFormatting } from "../documentStyles";
import {
  editableParagraphs,
  editParagraphs,
  type ParagraphSpot,
  paragraphPPr,
  selectedParagraphs,
} from "../paragraphEdits";

/**
 * The alignment this paragraph is rendered with.
 * If the paragraph wrote down no value it is the one the style gives, and failing that, left.
 */
function alignOf(node: PMNode): ParagraphAlign {
  return toParagraphFormat(node.attrs.format)?.align ?? "left";
}

/**
 * Sets the alignment of the selected paragraphs.
 * Exactly one alignment is always on, so there is no command to clear it (the same as Word).
 */
export function setParagraphAlign(align: ParagraphAlign): Command {
  return (state, dispatch) =>
    editParagraphs(state, dispatch, (node) =>
      // A paragraph already rendered with that alignment is left untouched, so its original XML survives
      alignOf(node) === align
        ? null
        : withParagraphAlign(paragraphPPr(node), align)
    );
}

/**
 * Whether there is a paragraph to align at all, which is what the alignment menu is drawn from.
 *
 * The four alignments are one choice, so no single one of them answers for the menu: the one the
 * selection already wears would report that it has nothing to do. What settles it is whether the
 * lock leaves any of the selected paragraphs open (`editor/paragraphEdits`), the same question
 * `canSetLineSpacing` asks for the menu beside it.
 */
export function canSetParagraphAlign(state: EditorState): boolean {
  return editableParagraphs(state).length > 0;
}

/** The alignment of the selected paragraphs */
export type ActiveParagraphAlign =
  /** They are all drawn with the same one */
  | { kind: "shared"; align: ParagraphAlign }
  /** They are drawn with several different ones, or the selection holds no paragraph at all */
  | { kind: "mixed" };

export function activeParagraphAlign(state: EditorState): ActiveParagraphAlign {
  const aligns = selectedParagraphs(state).map(({ node }) => alignOf(node));
  const first = aligns[0];
  return first !== undefined && aligns.every((align) => align === first)
    ? { kind: "shared", align: first }
    : { kind: "mixed" };
}

/** One run mark to be swapped in over the text it already covers */
interface MarkChange {
  from: number;
  to: number;
  mark: Mark;
}

/**
 * The run marks inside the paragraph, with the display values read again under the new style.
 *
 * A mark's display values were baked with the old style laid underneath the run's own
 * formatting, so leaving them alone would keep drawing the style the paragraph no longer wears.
 * Reading them again the way import does keeps the screen the same as it would be after saving
 * this file and reopening it. Every piece of text in the paragraph wears the style, the same as
 * Word, so text typed in the editor is marked here too rather than waiting for the next reopen.
 *
 * This edit leaves marks inside locked controls alone. The display deriver subsequently refreshes
 * their existing run marks under its source-preserving pass, so no content edit needs a lock bypass.
 */
function restyledMarks(
  spot: ParagraphSpot,
  paragraph: ResolvedParagraph,
  context: FormattingContext
): MarkChange[] {
  const changes: MarkChange[] = [];
  spot.node.forEach((child, offset) => {
    if (lockedMarkOf(child)) return;
    const mark = runMarkUnder(
      child.marks.find((entry) => entry.type === docxSchema.marks.run) ?? null,
      child.isText,
      paragraph,
      context
    );
    if (!mark) return;
    const from = spot.pos + 1 + offset;
    changes.push({ from, to: from + child.nodeSize, mark });
  });
  return changes;
}

/** One paragraph pointed at the new style, together with the marks inside it */
interface StyleChange {
  spot: ParagraphSpot;
  props: ParagraphProps;
  attrs: ParagraphAttrs;
  marks: MarkChange[];
}

/**
 * Writes the changes out in a single transaction.
 * Only the node attributes and the marks change, so no position moves and the selection
 * stays where it was.
 */
function writeStyleChanges(
  state: EditorState,
  changed: readonly StyleChange[]
): Transaction {
  const tr = state.tr;
  for (const { spot, props, attrs, marks } of changed) {
    tr.setNodeMarkup(tr.mapping.map(spot.pos), undefined, {
      ...spot.node.attrs,
      pPr: props.pPr,
      ...attrs,
    });
    for (const change of marks) {
      tr.addMark(change.from, change.to, change.mark);
    }
  }
  return tr;
}

/**
 * Points the selected paragraphs at one of the styles the document defines.
 * A null id takes the `w:pStyle` away, which is how a paragraph wears the default style.
 */
export function setParagraphStyle(styleId: string | null): Command {
  return (state, dispatch) => {
    const context = documentFormatting(state);
    // This writer is its own, so it leaves the locked paragraphs out itself, exactly as
    // `editParagraphs` does for every other paragraph edit
    const changed = editableParagraphs(state).flatMap((spot) => {
      const pPr = paragraphPPr(spot.node);
      // A paragraph already pointing at that style is left untouched, so its original XML survives
      if (styleIdOf(pPr) === styleId) return [];
      const props = withParagraphStyle(pPr, styleId);
      if (!props) return [];
      // The text takes the values of the style the paragraph now wears, the default one where the name was cleared
      const paragraph = resolveParagraph(props.pPr, context);
      return [
        {
          spot,
          props,
          attrs: paragraphAttrsOf(paragraph),
          marks: restyledMarks(spot, paragraph, context),
        },
      ];
    });
    if (changed.length === 0) return false;
    if (dispatch) dispatch(writeStyleChanges(state, changed).scrollIntoView());
    return true;
  };
}

/** The style the selected paragraphs point at */
export type ActiveParagraphStyle =
  /** They all point at the same one. A null id is a paragraph wearing the default style */
  | { kind: "shared"; styleId: string | null }
  /** They point at several different ones */
  | { kind: "mixed" }
  /** The selection holds no paragraph at all, so there is no style to show or to change */
  | { kind: "none" };

export function activeParagraphStyle(state: EditorState): ActiveParagraphStyle {
  const ids = selectedParagraphs(state).map(({ node }) =>
    styleIdOf(paragraphPPr(node))
  );
  if (ids.length === 0) return { kind: "none" };
  const first = ids[0] ?? null;
  return ids.every((id) => id === first)
    ? { kind: "shared", styleId: first }
    : { kind: "mixed" };
}
