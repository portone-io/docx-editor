/**
 * List editing commands.
 *
 * All they do is change a paragraph's place in a list (its numbering id and level).
 * Editing the formatting fragments belongs to `docx/paraProps`, and the list's appearance
 * belongs to `numbering/`.
 * Starting a list records its definition on the document node in the same transaction, and the
 * export writes that definition into numbering.xml.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import {
  type IndentChange,
  type ListChange,
  withLeftIndent,
  withListNumbering,
} from "../../docx/paraProps";
import { type NumberingRef, toParagraphFormat } from "../../model/format";
import {
  listsWorn,
  NEW_LISTS_ATTR,
  type NewLists,
  newListsValue,
} from "../../numbering/listRegistry";
import {
  allocateList,
  type ListKind,
  listFor,
  MAX_ILVL,
  templateList,
} from "../../numbering/listTemplate";
import type { Numbering } from "../../numbering/parseNumbering";
import { docxSchema } from "../../schema";
import { documentFormatting } from "../documentStyles";
import {
  editableParagraphs,
  editParagraphs,
  type ParagraphEditExtras,
  paragraphPPr,
  selectedParagraphs,
} from "../paragraphEdits";
import {
  canStartNewList,
  documentNumbering,
} from "../plugins/numberingDecorations";

/** This paragraph's place in a list. Null when it is not a list */
export function listRefOf(node: PMNode): NumberingRef | null {
  return toParagraphFormat(node.attrs.format)?.numbering ?? null;
}

/** The rule that decides what to change for each paragraph. Null skips that paragraph */
type ChangePlan = (node: PMNode) => ListChange | null;

/**
 * Edits the selected list paragraphs in a single transaction, and records the register the edit
 * leaves behind. A list nothing wears once the edit is in is not one the document has, so its
 * definition goes with it.
 */
function changeParagraphs(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  plan: ChangePlan,
  extras: ParagraphEditExtras = {
    alongside: pruneLists(documentNumbering(state).added),
  }
): boolean {
  return editParagraphs(
    state,
    dispatch,
    (node) => {
      const change = plan(node);
      return change && withListNumbering(paragraphPPr(node), change);
    },
    extras
  );
}

/**
 * What moving to another level does to the indentation.
 *
 * A paragraph whose own formatting fixes its left indentation is drawn at that indentation, so
 * it takes the new level's value written down to move it.
 * A paragraph that fixes none is drawn at the indentation its level specifies (overlaid by
 * `numberingDecorations`), so the level change alone moves it and nothing is written down.
 */
function indentForLevel(
  numbering: Numbering,
  node: PMNode,
  ref: NumberingRef
): IndentChange {
  const format = toParagraphFormat(node.attrs.format);
  if (
    format?.indentStartPt === undefined &&
    format?.indentLeftPt === undefined
  ) {
    return { kind: "keep" };
  }
  const indent = listFor(numbering, ref.numId)?.levels.get(ref.ilvl)?.indent;
  return indent ? { kind: "level", indent } : { kind: "keep" };
}

/**
 * The change that moves a list paragraph `delta` levels.
 * Null when the paragraph is not in a list or already sits at the first or the last level.
 * The indent buttons reuse this, so moving a level is the same edit however it is asked for.
 */
export function listLevelChange(
  numbering: Numbering,
  node: PMNode,
  delta: 1 | -1
): ListChange | null {
  const ref = listRefOf(node);
  if (!ref) return null;
  const ilvl = ref.ilvl + delta;
  if (ilvl < 0 || ilvl > MAX_ILVL) return null;
  const next = { numId: ref.numId, ilvl };
  return { numbering: next, indent: indentForLevel(numbering, node, next) };
}

/** Shifts a list paragraph one level. At the first and last levels it shifts no further */
function levelShift(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const numbering = documentNumbering(state);
    return changeParagraphs(state, dispatch, (node) =>
      listLevelChange(numbering, node, delta)
    );
  };
}

/** Moves a list item one level deeper, which is not the paragraph's own indent. What Tab does */
export const increaseListLevel: Command = levelShift(1);

/** Moves a list item one level back up, which is not the paragraph's own indent. What Shift+Tab does */
export const decreaseListLevel: Command = levelShift(-1);

/** Every numbering id a paragraph of this document wears */
export function numIdsIn(doc: PMNode): Set<number> {
  const worn = new Set<number>();
  doc.descendants((node) => {
    if (node.type !== docxSchema.nodes.paragraph) return true;
    const ref = listRefOf(node);
    if (ref) worn.add(ref.numId);
    return false;
  });
  return worn;
}

/** Writes down the register this edit leaves, which is `lists` minus what it left nothing wearing */
function recordLists(tr: Transaction, lists: NewLists): void {
  tr.setDocAttribute(
    NEW_LISTS_ATTR,
    newListsValue(listsWorn(lists, numIdsIn(tr.doc)))
  );
}

/**
 * Takes the register down to the lists the document still holds, and leaves the document node
 * alone when the edit orphaned none.
 */
function pruneLists(lists: NewLists): (tr: Transaction) => void {
  return (tr) => {
    if (listsWorn(lists, numIdsIn(tr.doc)) !== lists) recordLists(tr, lists);
  };
}

/**
 * Turns the selected paragraphs into a new list.
 * Paragraphs selected together share one numbering id, so they become a single continuous list.
 * Changing the list kind goes down this same path: taking a new numbering id is what decides
 * the new appearance.
 *
 * The definition the list is started with is registered on the document node in the same
 * transaction that puts the paragraphs in it, so what is drawn, what is exported and what undo
 * takes back are one and the same value rather than three readings of the number it was given.
 *
 * A new list can only be written out if its definition has a numbering part to go into, whether
 * the document arrived with one or the export writes it. Where neither is open to it we do not
 * start a list, blocking up front an edit that would be blocked at export time.
 *
 * The indentation of the paragraph is not touched. The one the list level specifies is drawn on
 * screen only (`numberingDecorations`) and goes out in the definition, so joining a list writes
 * down nothing that leaving it would have to take back.
 */
function startList(kind: ListKind): Command {
  return (state, dispatch) => {
    if (!canStartNewList(state)) return false;
    const started = allocateList(
      documentNumbering(state),
      numIdsIn(state.doc),
      templateList(kind)
    );
    return changeParagraphs(
      state,
      dispatch,
      (node) => ({
        numbering: { numId: started.numId, ilvl: listRefOf(node)?.ilvl ?? 0 },
        indent: { kind: "keep" },
      }),
      {
        formatting: {
          ...documentFormatting(state),
          numbering: started.numbering,
        },
        alongside: (tr) => recordLists(tr, started.numbering.added),
      }
    );
  };
}

/**
 * Takes the selected paragraphs out of the list.
 * Only the hanging indentation the marker sat in goes with it; an indentation the paragraph had
 * of its own stays, putting it back where it was before it joined the list.
 */
export const removeFromList: Command = (state, dispatch) =>
  changeParagraphs(state, dispatch, (node) =>
    listRefOf(node)
      ? { numbering: null, indent: { kind: "clearHanging" } }
      : null
  );

/**
 * The kind of this list position, which the number format of its level decides.
 * Null where nothing defines the list, which is a position whose kind cannot be told and so is
 * neither of the two buttons.
 */
function kindOf(numbering: Numbering, ref: NumberingRef): ListKind | null {
  const level = listFor(numbering, ref.numId)?.levels.get(ref.ilvl);
  if (!level) return null;
  return level.format === "bullet" ? "bullet" : "numbered";
}

/**
 * The kind of list the selected paragraphs belong to. Null when they are not a list, when nothing
 * defines the list they are in, or when the kinds are mixed.
 * It reads the paragraphs a lock leaves open, the same ones `editParagraphs` writes to: counting a
 * locked paragraph that stays out of the list would keep the answer mixed, so the button would make
 * a list a second time instead of taking it off.
 */
export function activeListKind(state: EditorState): ListKind | null {
  const numbering = documentNumbering(state);
  const kinds = editableParagraphs(state).map(({ node }) => {
    const ref = listRefOf(node);
    return ref === null ? null : kindOf(numbering, ref);
  });
  const first = kinds[0] ?? null;
  return kinds.every((kind) => kind === first) ? first : null;
}

/**
 * What the list button does.
 * If it is already that kind, the paragraphs leave the list; if it is a different kind or not
 * a list at all, they become that kind (the same as Word).
 */
function toggleList(kind: ListKind): Command {
  return (state, dispatch) =>
    activeListKind(state) === kind
      ? removeFromList(state, dispatch)
      : startList(kind)(state, dispatch);
}

export const toggleNumberedList: Command = toggleList("numbered");

export const toggleBulletList: Command = toggleList("bullet");

/**
 * Whether the selected paragraphs hold a list item, which is a fact about the document and so is
 * answered whatever the lock says, as `activeListKind` beside it is.
 * What a level shift would actually move is the level-shift command's own answer, the lock already
 * part of it, and that is what a button runs on.
 */
export function isInList(state: EditorState): boolean {
  return selectedParagraphs(state).some(
    (spot) => listRefOf(spot.node) !== null
  );
}

/**
 * Takes the paragraph out of the list and leaves it at the true start of the line.
 *
 * The marker slot goes, and so does the whole left indent, so nothing of the paragraph's list
 * life is left to hold the caret away from the line start.
 * A left indent the document gave the paragraph before it ever joined the list goes with it:
 * `w:ind` does not record where a value came from, and abandoning an empty item is a request for
 * a fresh line rather than for the paragraph's earlier shape. `removeFromList`, what the list
 * button does, is the path that puts a paragraph back the way it was.
 * A first-line indent is not part of a list - only a hanging indent is a marker slot - so that
 * one is left alone, as is the right indent.
 */
const leaveListAtLineStart: Command = (state, dispatch) =>
  editParagraphs(
    state,
    dispatch,
    (node) => {
      if (!listRefOf(node)) return null;
      const unlisted = withListNumbering(paragraphPPr(node), {
        numbering: null,
        indent: { kind: "clearHanging" },
      });
      return unlisted && withLeftIndent(unlisted.pPr, 0);
    },
    { alongside: pruneLists(documentNumbering(state).added) }
  );

/**
 * Pressing Enter in an empty list item does not open a new item.
 *
 * A deeper item comes up one level and stays in the list, and only an item already at the first
 * level leaves it, so a nested list is climbed one press at a time (the same as Word).
 * The paragraph itself stays where it is, which leaves the items after it in the list with their
 * numbering slots untouched; they simply stop counting the item that was given up.
 */
export const leaveEmptyListItem: Command = (state, dispatch) => {
  const { empty, $from } = state.selection;
  const paragraph = $from.parent;
  if (!empty || paragraph.type !== docxSchema.nodes.paragraph) return false;
  if (paragraph.content.size > 0) return false;
  const ref = listRefOf(paragraph);
  if (!ref) return false;
  return ref.ilvl > 0
    ? decreaseListLevel(state, dispatch)
    : leaveListAtLineStart(state, dispatch);
};
