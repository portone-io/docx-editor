import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import { DocAttrStep } from "prosemirror-transform";
import { noteLabelSignature, noteLabelsIn } from "../../docx/notes/numbering";
import { sectionBreakOf } from "../../docx/sections";
import { displayOnly } from "../../schema/displayDerivation";
import { stepReaches } from "../../schema/guards";
import { documentOf, type EditorDocument } from "../editorDocument";

export type NoteLabeller = typeof noteLabelsIn;

/** A node a label is worked out from: a note reference, or a paragraph that ends a section */
function countsNotes(node: PMNode): boolean {
  if (node.type.name === "noteReference") return true;
  const pPr: unknown = node.attrs.pPr;
  return (
    node.type.name === "paragraph" &&
    typeof pPr === "string" &&
    sectionBreakOf(pPr) !== null
  );
}

/**
 * Whether the transaction reaches something a label is worked out from.
 *
 * The body's own section stands on the document node, whose other attrs - the side stories among
 * them - change far more often, so a document attr step reaches a label only when it is that one.
 */
function reachesLabels(tr: Transaction): boolean {
  return tr.steps.some((step, index) =>
    step instanceof DocAttrStep
      ? step.attr === "sectPr"
      : stepReaches(
          step,
          tr.docs[index],
          tr.docs[index + 1] ?? tr.doc,
          countsNotes
        )
  );
}

function labelsMayHaveMoved(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState
): boolean {
  return (
    transactions.some(reachesLabels) &&
    noteLabelSignature(oldState.doc) !== noteLabelSignature(newState.doc)
  );
}

function numberingReplaced(
  before: EditorDocument,
  after: EditorDocument
): boolean {
  return (
    before.noteNumbering !== after.noteNumbering ||
    before.specialNotes !== after.specialNotes
  );
}

/**
 * Labels the note references again after a change that moves one, moves a section break, or
 * replaces the numbering the snapshot holds.
 *
 * A label is a display value of an inline node, which the display derivation walk does not reach,
 * so it is written here under the same pass: a relabel goes through a lock around the reference,
 * and one undo takes it back with the edit that caused it. The relabel after a snapshot change
 * follows no edit and goes to the history not at all.
 */
export function noteNumbering(labelsIn: NoteLabeller = noteLabelsIn): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      const document = documentOf(newState);
      const replaced = numberingReplaced(documentOf(oldState), document);
      if (!replaced && !labelsMayHaveMoved(transactions, oldState, newState)) {
        return null;
      }
      const labels = labelsIn(
        newState.doc,
        document.noteNumbering,
        document.specialNotes
      );
      const tr = newState.tr;
      for (const [pos, label] of labels) {
        if (tr.doc.nodeAt(pos)?.attrs.label !== label) {
          tr.setNodeAttribute(pos, "label", label);
        }
      }
      if (!tr.docChanged) return null;
      tr.setMeta(displayOnly, true);
      if (replaced) tr.setMeta("addToHistory", false);
      return tr;
    },
  });
}
