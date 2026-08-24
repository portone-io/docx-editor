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
import { layerRunFormat, styleIdOf } from "../../docx/formatting";
import {
  type ParagraphProps,
  withParagraphAlign,
  withParagraphStyle,
} from "../../docx/paraProps";
import { readRunProps } from "../../docx/runProps";
import {
  type ParagraphAlign,
  type RunFormat,
  toParagraphFormat,
} from "../../model/format";
import { docxSchema } from "../../schema";
import { lockedMarkOf } from "../../schema/locks";
import { defaultParagraphStyleId, documentStyles } from "../documentStyles";
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
    editParagraphs(state, dispatch, (node, styles, defaultStyleId) =>
      // A paragraph already rendered with that alignment is left untouched, so its original XML survives
      alignOf(node) === align
        ? null
        : withParagraphAlign(paragraphPPr(node), align, styles, defaultStyleId)
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

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/** One run mark to be swapped in over the text it already covers */
interface MarkChange {
  from: number;
  to: number;
  mark: Mark;
}

/**
 * The run mark one inline child wears under the new style, or null when it needs none.
 *
 * A child already carrying a mark keeps the run formatting it wrote down and only has its
 * display values read again, because they were baked with the old style laid underneath.
 * Text typed in the editor carries no mark at all, and gets one holding nothing but the
 * values the style gives, which is the very mark import builds for a bare `w:r`.
 * A style laying down no run formatting leaves nothing to attach.
 */
function restyledMark(child: PMNode, style: RunFormat): Mark | null {
  const mark = child.marks.find((entry) => entry.type === docxSchema.marks.run);
  if (mark) {
    const format = layerRunFormat(style, readRunProps(text(mark.attrs.rPr)));
    return mark.type.create({ ...mark.attrs, format });
  }
  if (!child.isText) return null;
  const format = layerRunFormat(style, null);
  if (format === null) return null;
  return docxSchema.marks.run.create({ rPr: null, rAttrs: null, format });
}

/**
 * The run marks inside the paragraph, with the display values read again under the new style.
 *
 * A mark's display values were baked with the old style laid underneath the run's own
 * formatting, so leaving them alone would keep drawing the style the paragraph no longer wears.
 * Reading them again the way import does keeps the screen the same as it would be after saving
 * this file and reopening it. Every piece of text in the paragraph wears the style, the same as
 * Word, so text typed in the editor is marked here too rather than waiting for the next reopen.
 * Such a mark writes no `w:rPr` of its own, so it changes nothing on the way out.
 * An inline that is not text (an image, say) and carries no mark is left as it is.
 *
 * Text inside a locked control is left alone as well. The guard turns down the whole transaction
 * over it, so the other selected paragraphs would go unstyled with it; the paragraph still comes
 * to point at the new style, the same as alignment and indent already do.
 */
function restyledMarks(spot: ParagraphSpot, style: RunFormat): MarkChange[] {
  const changes: MarkChange[] = [];
  spot.node.forEach((child, offset) => {
    if (lockedMarkOf(child)) return;
    const mark = restyledMark(child, style);
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
  for (const { spot, props, marks } of changed) {
    tr.setNodeMarkup(tr.mapping.map(spot.pos), undefined, {
      ...spot.node.attrs,
      pPr: props.pPr,
      format: props.format,
      styleRun: props.styleRun,
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
    const styles = documentStyles(state);
    const defaultStyleId = defaultParagraphStyleId(state);
    // This writer is its own, so it leaves the locked paragraphs out itself, exactly as
    // `editParagraphs` does for every other paragraph edit
    const changed = editableParagraphs(state).flatMap((spot) => {
      const pPr = paragraphPPr(spot.node);
      // A paragraph already pointing at that style is left untouched, so its original XML survives
      if (styleIdOf(pPr) === styleId) return [];
      const props = withParagraphStyle(pPr, styleId, styles, defaultStyleId);
      if (!props) return [];
      // The text takes the values of the style the paragraph now wears, the default one where the name was cleared
      return [
        { spot, props, marks: restyledMarks(spot, props.styleRun ?? {}) },
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
