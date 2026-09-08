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
import { paragraphAttrsFor } from "../docx/formatting";
import type { ParagraphProps } from "../docx/paraProps";
import { docxSchema } from "../schema";
import { editShut } from "../schema/guards";
import { documentFormatting } from "./documentStyles";
import { paragraphPlacementAt } from "./paragraphPlacement";

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
 * The selected paragraphs the guards leave editable, which is what every paragraph edit works on.
 *
 * A shut paragraph is left out rather than the whole edit refused, the policy character formatting
 * already follows (`schema/guards`): the guard turns down a whole transaction, so asking for it
 * would leave the rest of the selection unedited too. A selection the guards leave nothing of edits
 * nothing, and the command reports that of its own accord, which is the disabled state of the
 * control that runs it.
 *
 * A paragraph is rewritten around its content, which is the block intent: only the paragraphs of a
 * locked cell are shut, and a paragraph merely holding a locked control keeps its own alignment,
 * indent and style. A protection that shuts the body shuts every one of them.
 */
export function editableParagraphs(state: EditorState): ParagraphSpot[] {
  return selectedParagraphs(state).filter(
    (spot) => !editShut(state, { kind: "block", at: spot.pos })
  );
}

/** The original formatting XML this paragraph holds. Null when there is none */
export function paragraphPPr(node: PMNode): string | null {
  const pPr: unknown = node.attrs.pPr;
  return typeof pPr === "string" ? pPr : null;
}

/** How to edit a single paragraph. Null skips that paragraph */
export type ParagraphSurgery = (node: PMNode) => ParagraphProps | null;

interface PlannedChange {
  spot: ParagraphSpot;
  props: ParagraphProps;
}

/**
 * What an edit records beside the paragraphs themselves, such as the definition of the list they
 * are joining. It is written into the same transaction, so undo takes the two back together.
 */
export type AlongsideParagraphs = (tr: Transaction) => void;

function writeChanges(
  state: EditorState,
  changed: readonly PlannedChange[],
  alongside: AlongsideParagraphs | undefined
): Transaction {
  const tr = state.tr;
  const formatting = documentFormatting(state);
  for (const { spot, props } of changed) {
    tr.setNodeMarkup(tr.mapping.map(spot.pos), undefined, {
      ...spot.node.attrs,
      pPr: props.pPr,
      ...paragraphAttrsFor(
        props.pPr,
        formatting,
        paragraphPlacementAt(state.doc, spot.pos)
      ),
    });
  }
  alongside?.(tr);
  return tr;
}

/** Edits the selected paragraphs in a single transaction. Does nothing when there is nothing to edit */
export function editParagraphs(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  surgery: ParagraphSurgery,
  alongside?: AlongsideParagraphs
): boolean {
  const changed = editableParagraphs(state).flatMap((spot) => {
    const props = surgery(spot.node);
    return props ? [{ spot, props }] : [];
  });
  if (changed.length === 0) return false;
  if (dispatch) {
    dispatch(writeChanges(state, changed, alongside).scrollIntoView());
  }
  return true;
}
