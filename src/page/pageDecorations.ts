/**
 * Decorations that move content to the next page: the push above a block, and whatever the block's
 * own kind draws the cuts inside it as (`page/kinds`).
 *
 * The kinds the editor was built with are held here, so the measurement and the decorations read
 * one registry and a block is measured by the same kind that draws it. The demand sources
 * (`page/demands`) the measurement asks about every block are held beside them.
 * The document model is left untouched, so no trace of any of it is left in the exported XML or
 * in the edit history.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  Plugin,
  PluginKey,
  type Transaction,
} from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { editorAttributes } from "../styles/classNames";
import { type BlockKind, blockKindFor, type PageCut } from "./blockKinds";
import { DEFAULT_DEMAND_SOURCES, type DemandSource } from "./demands";
import { DEFAULT_BLOCK_KINDS } from "./kinds";
import type { BlockPush } from "./pageLayout";

/** Everything one measurement has to say about the page (`page/pageLayout`) */
export interface PageMarksInput {
  pushes: readonly BlockPush[];
  /**
   * Where the layout parted a block, keyed by the position the continued piece starts at.
   *
   * The layout counts within the block it measured, and an edit anywhere before a cut shifts that
   * count, so what is kept here is a position: a position maps through an edit, an ordinal does
   * not.
   */
  cuts: readonly PageCut[];
}

interface PageMarks extends PageMarksInput {
  kinds: readonly BlockKind[];
  sources: readonly DemandSource[];
  decorations: DecorationSet;
}

/** What the editor was built with, which every set of marks carries on to the next */
type PageRegistry = Pick<PageMarks, "kinds" | "sources">;

const marksKey = new PluginKey<PageMarks>("docxPageDecorations");

/**
 * A paragraph may already carry a `margin-top` of its own.
 * ProseMirror clears the property written here when the decoration is taken away, and
 * layering on the logical property that means the same thing leaves the value the
 * paragraph originally had intact.
 * The later declaration wins, so while the push is in place this is the value used.
 */
function pushStyle(marginTop: number): string {
  return `margin-block-start:${marginTop}px`;
}

/** The cuts of each top-level block, keyed by the position that block starts at */
function cutsByBlock(
  doc: PMNode,
  cuts: readonly PageCut[]
): Map<number, PageCut[]> {
  const byBlock = new Map<number, PageCut[]>();
  for (const cut of cuts) {
    const $at = doc.resolve(cut.at);
    if ($at.depth === 0) continue;
    const blockPos = $at.before(1);
    const found = byBlock.get(blockPos);
    if (found) found.push(cut);
    else byBlock.set(blockPos, [cut]);
  }
  return byBlock;
}

/**
 * Every block is handed to its kind whether the layout gave it a cut or not, because a kind draws
 * what its own measurement reads: a paragraph's page break carries an empty space until a
 * measurement says otherwise, and the measurement reads where the break stands off that very
 * element.
 */
function decorationsFor(
  doc: PMNode,
  kinds: readonly BlockKind[],
  pushes: readonly BlockPush[],
  cuts: readonly PageCut[]
): DecorationSet {
  const byPos = new Map(pushes.map((push) => [push.pos, push]));
  const byBlock = cutsByBlock(doc, cuts);
  const decorations: Decoration[] = [];
  doc.forEach((node, offset) => {
    const push = byPos.get(offset);
    if (push) {
      decorations.push(
        Decoration.node(offset, offset + node.nodeSize, {
          style: pushStyle(push.marginTop),
          [editorAttributes.pagePush]: `${push.push}`,
        })
      );
    }
    blockKindFor(kinds, node).decorate(
      offset,
      node,
      byBlock.get(offset) ?? [],
      decorations
    );
  });
  return DecorationSet.create(doc, decorations);
}

function marksFor(
  doc: PMNode,
  { kinds, sources }: PageRegistry,
  pushes: readonly BlockPush[],
  cuts: readonly PageCut[]
): PageMarks {
  return {
    kinds,
    sources,
    pushes,
    cuts,
    decorations: decorationsFor(doc, kinds, pushes, cuts),
  };
}

/** Whether the block the cut stands in still holds a place its kind can cut at */
function stillCuts(
  kinds: readonly BlockKind[],
  doc: PMNode,
  at: number
): boolean {
  const $at = doc.resolve(at);
  if ($at.depth === 0) return false;
  const block = doc.nodeAt($at.before(1));
  return block !== null && blockKindFor(kinds, block).holdsCut(doc, at);
}

/**
 * Whether the node the cut named is still there after an edit that replaced its opening token,
 * as `setNodeMarkup` does to a row. Such an edit maps the node's own position as deleted, and
 * what tells it apart from a node that really went away is the content: this one still has its
 * content boundary just inside it.
 */
function retainsContent(tr: Transaction, at: number, mapped: number): boolean {
  const node = tr.before.nodeAt(at);
  if (!node || node.isLeaf) return false;
  const content = tr.mapping.mapResult(at + 1, -1);
  return !content.deletedAcross && content.pos === mapped + 1;
}

function samePageMarks(a: PageMarksInput, b: PageMarksInput): boolean {
  return (
    a.pushes.length === b.pushes.length &&
    a.cuts.length === b.cuts.length &&
    a.pushes.every((push, index) => {
      const other = b.pushes[index];
      return (
        other !== undefined &&
        other.pos === push.pos &&
        other.marginTop === push.marginTop &&
        other.push === push.push
      );
    }) &&
    a.cuts.every((cut, index) => {
      const other = b.cuts[index];
      return (
        other !== undefined &&
        other.at === cut.at &&
        other.height === cut.height
      );
    })
  );
}

export function pageDecorations(
  kinds: readonly BlockKind[] = DEFAULT_BLOCK_KINDS,
  sources: readonly DemandSource[] = DEFAULT_DEMAND_SOURCES
): Plugin<PageMarks> {
  return new Plugin<PageMarks>({
    key: marksKey,
    state: {
      init: (_config, state) => marksFor(state.doc, { kinds, sources }, [], []),
      apply(tr, value) {
        const next = tr.getMeta(marksKey);
        if (next) return next;
        if (!tr.docChanged) return value;
        // Positions shift when the text changes, and a break the edit has just put in has no
        // space yet, so the marks are laid out again over the new document. A cut whose piece
        // the edit took away goes with it, rather than landing on whatever the mapping now
        // points at
        return marksFor(
          tr.doc,
          value,
          value.pushes.map((push) => ({
            ...push,
            pos: tr.mapping.map(push.pos),
          })),
          value.cuts.flatMap((cut) => {
            const mapped = tr.mapping.mapResult(cut.at, 1);
            return !stillCuts(value.kinds, tr.doc, mapped.pos) ||
              (mapped.deleted && !retainsContent(tr, cut.at, mapped.pos))
              ? []
              : [{ ...cut, at: mapped.pos }];
          })
        );
      },
    },
    props: {
      decorations: (state) => marksKey.getState(state)?.decorations,
    },
  });
}

/** The kinds the editor was built with, which the measurement reads to match them */
export function blockKindsOf(state: EditorState): readonly BlockKind[] {
  return marksKey.getState(state)?.kinds ?? DEFAULT_BLOCK_KINDS;
}

/** The demand sources the editor was built with, which the measurement asks about every block */
export function demandSourcesOf(state: EditorState): readonly DemandSource[] {
  return marksKey.getState(state)?.sources ?? DEFAULT_DEMAND_SOURCES;
}

/** The cuts the sheet is drawn with, whose spaces the measurement takes back off a demand's place */
export function pageCutsOf(state: EditorState): readonly PageCut[] {
  return marksKey.getState(state)?.cuts ?? [];
}

function current(view: EditorView): PageMarksInput {
  return marksKey.getState(view.state) ?? { pushes: [], cuts: [] };
}

/** One transaction, or none when the same marks are already applied */
export function setPageMarks(view: EditorView, next: PageMarksInput): void {
  if (samePageMarks(current(view), next)) return;
  view.dispatch(
    view.state.tr
      .setMeta(
        marksKey,
        marksFor(
          view.state.doc,
          {
            kinds: blockKindsOf(view.state),
            sources: demandSourcesOf(view.state),
          },
          next.pushes,
          next.cuts
        )
      )
      .setMeta("addToHistory", false)
  );
}
