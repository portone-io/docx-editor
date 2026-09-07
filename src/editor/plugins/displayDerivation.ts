/**
 * Keeps every display value of the document current: the ones an edit moved, after that edit, and
 * every one at once when the document snapshot the values are resolved against is replaced.
 *
 * The derivers are registered here and run through the one walk `schema/displayDerivation`
 * defines. What a node maps back to is settled here as well, from the mapping of the transactions
 * that changed the document: a node whose content survived those transactions is the node that
 * stood where its opening maps back to, and one whose content did not is new. A node rewritten
 * where it stands (`setNodeMarkup`) therefore still maps back to itself, which is what lets a
 * table resized or a paragraph re-aligned be told from a table or a paragraph put in.
 *
 * The transaction this appends carries the display-only pass (`schema/guards`), since it changes
 * nothing a lock or a protection answers for, and goes to the history as part of the edit it
 * follows, so that one undo takes the edit and its values back together. The re-derivation after
 * a snapshot change follows no edit and goes to the history not at all: the values follow the
 * snapshot, and undoing whatever replaced the snapshot derives them again.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import { Mapping, Transform } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";
import {
  type DisplayDeriver,
  deriveDisplay,
  displayOnly,
} from "../../schema/displayDerivation";
import { documentOf, type EditorDocument } from "../editorDocument";
import { paragraphDisplay } from "./paragraphDisplay";
import { tableDisplay } from "./tableDisplay";

/** What a deriver of this editor resolves against */
export interface DerivationContext {
  readonly document: EditorDocument;
  /**
   * Where the composition the browser holds open stands, which a deriver leaves alone: rewriting
   * the node it stands in redraws it, and a redraw under an open composition is what the browser
   * answers by taking the composition down. Null while nothing is being composed, and while every
   * value is being worked out again from nothing, when nothing may be left standing.
   */
  readonly composingAt: number | null;
}

export type DocumentDeriver = DisplayDeriver<DerivationContext>;

export const DISPLAY_DERIVERS: readonly DocumentDeriver[] = [
  tableDisplay,
  paragraphDisplay,
];

const nothingBefore = (): null => null;

/**
 * The document with every display value worked out against this snapshot, which is what a state
 * is built over: the values a document arrived with were worked out against whatever opened it.
 */
export function withDerivedDisplay(
  doc: PMNode,
  document: EditorDocument,
  derivers: readonly DocumentDeriver[] = DISPLAY_DERIVERS
): PMNode {
  const transform = new Transform(doc);
  deriveDisplay(
    transform,
    { document, composingAt: null },
    derivers,
    nothingBefore
  );
  return transform.doc;
}

/**
 * The node each node of the changed document was, in the document before the transactions.
 *
 * The opening of a node rewritten where it stands maps back as taken away, and so does the
 * opening of a node put in; what tells the two apart is that the first still holds what it held.
 * So the position just inside the node is what is asked about, and a node that kept it is the
 * node of the same type that stood where its opening maps back to.
 */
function previousIn(
  before: PMNode,
  transactions: readonly Transaction[]
): (node: PMNode, pos: number) => PMNode | null {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  const back = mapping.invert();
  return (node, pos) => {
    const inside = node.isLeaf ? pos : pos + 1;
    if (back.mapResult(inside, 1).deleted) return null;
    const was = before.nodeAt(back.map(pos, 1));
    return was !== null && was.type === node.type ? was : null;
  };
}

/**
 * Where the composition the browser has open stands.
 * Null while nothing is being composed, and while the editor stands outside a view at all.
 */
function composingAt(
  view: EditorView | null,
  state: EditorState
): number | null {
  return view?.composing === true ? state.selection.from : null;
}

/**
 * Derives the display values again after every change: for the nodes an edit moved, and for every
 * node when the snapshot was replaced.
 *
 * An appended transaction is handed nothing but states, so the view an open composition is read
 * from is the one the plugin is given when it mounts.
 */
export function displayDerivation(
  derivers: readonly DocumentDeriver[] = DISPLAY_DERIVERS
): Plugin {
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
    appendTransaction(transactions, oldState, newState) {
      const document = documentOf(newState);
      const replaced = document !== documentOf(oldState);
      if (
        !replaced &&
        !transactions.some((transaction) => transaction.docChanged)
      ) {
        return null;
      }
      const tr = newState.tr.setMeta(displayOnly, true);
      if (replaced) tr.setMeta("addToHistory", false);
      deriveDisplay(
        tr,
        {
          document,
          composingAt: replaced ? null : composingAt(live, newState),
        },
        derivers,
        replaced ? nothingBefore : previousIn(oldState.doc, transactions)
      );
      return tr.docChanged ? tr : null;
    },
  });
}
