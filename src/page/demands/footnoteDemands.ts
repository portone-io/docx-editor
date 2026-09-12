/**
 * The room a footnote reference asks of the page it lands on: its footnote, drawn at the foot of
 * that page (`ui/notes`).
 *
 * The sheet is measured on every layout pass and nearly every block holds no reference, so which
 * references a block holds is read off the block node and remembered against it. A node an edit
 * did not touch is the same node afterwards, so a pass reads nothing off the page for a block with
 * no reference and walks only the blocks an edit rebuilt.
 */

import type { Node as PMNode } from "prosemirror-model";
import { docxSchema } from "../../schema";
import { type NoteKind, type StoryKey, storyKey } from "../../schema/stories";
import type { MeasureTarget } from "../blockKinds";
import type { DemandSource, PageDemand } from "./index";

/** The band footnotes are kept in, which the layout's caller hands the heights of */
export const FOOTNOTE_BAND = "footnote";

/**
 * Where footnotes stand among the bands a page keeps (`DemandBand.order`). Word sets them at the
 * very foot of the body, under anything else the page holds there, so this is the last place
 */
export const FOOTNOTE_BAND_ORDER = 0;

/**
 * The kinds of note a page keeps room for at its foot.
 *
 * Endnotes stand after the last block of the document rather than beside the text that calls
 * them, so they are rows the layout lays last (`page/pageLayout`) and no page keeps a band for
 * them. Asking which kinds belong here rather than which one does not keeps a reference of a kind
 * this editor does not model yet from claiming a footnote's room: it would reserve the band's
 * overhead on its page and draw a rule over an empty strip, with nothing to say what went wrong.
 * These are not `EDITABLE_NOTE_KINDS`, which is about what an edit may rewrite and holds both.
 */
const BAND_NOTE_KINDS: readonly NoteKind[] = ["footnote"];

/** One footnote reference inside a block: where it stands from the block's content start */
interface HeldReference {
  readonly offset: number;
  readonly key: StoryKey;
}

const NO_REFERENCES: readonly HeldReference[] = [];
const NO_DEMANDS: readonly PageDemand[] = [];

const held = new WeakMap<PMNode, readonly HeldReference[]>();

function footnoteReferencesIn(block: PMNode): readonly HeldReference[] {
  const remembered = held.get(block);
  if (remembered) return remembered;
  const found: HeldReference[] = [];
  block.descendants((node, offset) => {
    if (node.type !== docxSchema.nodes.noteReference) return true;
    const id: unknown = node.attrs.id;
    const kind = BAND_NOTE_KINDS.find(
      (candidate) => candidate === node.attrs.kind
    );
    if (kind !== undefined && typeof id === "string") {
      found.push({ offset, key: storyKey(kind, id) });
    }
    return false;
  });
  const references = found.length === 0 ? NO_REFERENCES : found;
  held.set(block, references);
  return references;
}

export const footnoteDemands: DemandSource = {
  name: "footnote",

  demandsIn({ view, node, pos, sheetY, top }: MeasureTarget) {
    const references = footnoteReferencesIn(node);
    if (references.length === 0) return NO_DEMANDS;
    return references.flatMap(({ offset, key }): PageDemand[] => {
      const drawn = view.nodeDOM(pos + 1 + offset);
      if (!(drawn instanceof Element)) return [];
      return [
        {
          offset: sheetY(drawn.getBoundingClientRect().top) - top,
          id: key,
          band: FOOTNOTE_BAND,
        },
      ];
    });
  },
};
