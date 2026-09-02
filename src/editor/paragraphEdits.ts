/**
 * The shared path that swaps the formatting fragments of the selected paragraphs in a single
 * transaction.
 *
 * All this file takes on is which paragraphs to edit and how to write the new fragments in.
 * How to edit a fragment is decided by the caller (lists, alignment), and the XML surgery is
 * done by `docx/paraProps`.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { effectiveParagraphFormat, type StyleTable } from "../docx/formatting";
import type { ParagraphProps } from "../docx/paraProps";
import { docxSchema } from "../schema";
import { insideLockedCell } from "../schema/locks";
import { editsShut } from "../schema/protectionState";
import {
  defaultParagraphStyleId,
  documentParagraphFormatting,
  documentStyles,
} from "./documentStyles";

export interface ParagraphSpot {
  pos: number;
  node: PMNode;
}

/** The paragraphs the selection spans. Selections covering several table cells are handled too */
export function selectedParagraphs(state: EditorState): ParagraphSpot[] {
  const spots = new Map<number, PMNode>();
  for (const range of state.selection.ranges) {
    state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
      if (node.type !== docxSchema.nodes.paragraph) return true;
      spots.set(pos, node);
      return false;
    });
  }
  return Array.from(spots, ([pos, node]) => ({ pos, node }));
}

/**
 * The selected paragraphs a lock leaves editable, which is what every paragraph edit works on.
 *
 * A locked stretch is left out rather than the whole edit refused, the policy character formatting
 * already follows (`commands/formattingCommands`): the guard turns down a whole transaction, so
 * asking for it would leave the rest of the selection unedited too. A selection the lock leaves
 * nothing of edits nothing, and the command reports that of its own accord, which is the disabled
 * state of the control that runs it.
 *
 * Only the paragraphs of a locked cell are shut. A paragraph merely holding a locked control keeps
 * its own alignment, indent and style, which is what `insideLockedCell` says and
 * `rangeTouchesLocked` would not (`schema/locks`).
 *
 * A protection that shuts the body leaves no paragraph editable (`schema/protection`).
 */
export function editableParagraphs(state: EditorState): ParagraphSpot[] {
  if (editsShut(state)) return [];
  return selectedParagraphs(state).filter(
    (spot) => !insideLockedCell(state.doc, spot.pos)
  );
}

/** The original formatting XML this paragraph holds. Null when there is none */
export function paragraphPPr(node: PMNode): string | null {
  const pPr: unknown = node.attrs.pPr;
  return typeof pPr === "string" ? pPr : null;
}

/**
 * How to edit a single paragraph. Null skips that paragraph.
 * The style table and the document's default paragraph style come along, because the display values
 * of the edited fragment are read back with the style the paragraph wears laid underneath.
 */
export type ParagraphSurgery = (
  node: PMNode,
  styles: StyleTable,
  defaultStyleId: string | null
) => ParagraphProps | null;

interface PlannedChange {
  spot: ParagraphSpot;
  props: ParagraphProps;
}

function writeChanges(
  state: EditorState,
  changed: readonly PlannedChange[]
): Transaction {
  const tr = state.tr;
  const formatting = documentParagraphFormatting(state);
  for (const { spot, props } of changed) {
    tr.setNodeMarkup(tr.mapping.map(spot.pos), undefined, {
      ...spot.node.attrs,
      pPr: props.pPr,
      format: effectiveParagraphFormat(props.pPr, formatting),
      styleRun: props.styleRun,
    });
  }
  return tr;
}

/** Edits the selected paragraphs in a single transaction. Does nothing when there is nothing to edit */
export function editParagraphs(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  surgery: ParagraphSurgery
): boolean {
  const styles = documentStyles(state);
  const defaultStyleId = defaultParagraphStyleId(state);
  const changed = editableParagraphs(state).flatMap((spot) => {
    const props = surgery(spot.node, styles, defaultStyleId);
    return props ? [{ spot, props }] : [];
  });
  if (changed.length === 0) return false;
  if (dispatch) dispatch(writeChanges(state, changed).scrollIntoView());
  return true;
}
