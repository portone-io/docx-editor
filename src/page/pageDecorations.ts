/**
 * Decorations that move content to the next page: block pushes, page-break spaces, and table
 * continuation rows.
 *
 * The document model is left untouched, so no trace of them is left in
 * the exported XML or in the edit history.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { docxSchema, isPageBreak } from "../schema";
import { editorAttributes } from "../styles/classNames";
import type { PageCut } from "./blockKinds";
import type { BlockPush } from "./pageLayout";
import { columnCount, headerRowsOf } from "./tableMeasurements";

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
  decorations: DecorationSet;
}

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

/**
 * A block box around the `br`, so what follows the break is laid out below it and a height moves
 * it down by exactly that much. An empty one redraws the paragraph as it stood, except one holding
 * nothing but the break, which loses the line the `br` alone occupied.
 */
function spaceStyle(height: number): string {
  return `display:block;height:${height}px`;
}

function isPageBreakNode(node: PMNode | null | undefined): boolean {
  return (
    node?.type === docxSchema.nodes.hardBreak && isPageBreak(node.attrs.brAttrs)
  );
}

function isRow(node: PMNode | null | undefined): boolean {
  return node?.type.spec.tableRole === "row";
}

/** One page break inside a block, where it stands and how much room the `br` itself takes */
interface BreakAt {
  at: number;
  size: number;
}

/**
 * Every page break inside this block, in document order.
 * A space is put on the breaks of a top-level paragraph alone: inside a table cell it would grow
 * the cell rather than the page, so a break there is left to the whole-block rule in
 * `page/measureBlocks`, and none is found here.
 *
 * The measurement reads the same list to pair each space element it finds with the break it
 * belongs to (`page/measureBlocks`).
 */
export function pageBreaksIn(block: PMNode, blockPos: number): BreakAt[] {
  const found: BreakAt[] = [];
  if (block.type !== docxSchema.nodes.paragraph) return found;
  block.forEach((child, offset) => {
    if (isPageBreakNode(child)) {
      found.push({ at: blockPos + 1 + offset, size: child.nodeSize });
    }
  });
  return found;
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

function breakSpaces(
  block: PMNode,
  blockPos: number,
  cuts: readonly PageCut[],
  into: Decoration[]
): void {
  const heights = new Map(cuts.map((cut) => [cut.at, cut.height]));
  for (const { at, size } of pageBreaksIn(block, blockPos)) {
    const height = heights.get(at) ?? 0;
    into.push(
      Decoration.inline(at, at + size, {
        nodeName: "span",
        style: spaceStyle(height),
        [editorAttributes.pageBreakSpace]: `${height}`,
      })
    );
  }
}

function tableSpace(height: number, columns: number): HTMLElement {
  const row = document.createElement("tr");
  row.setAttribute(editorAttributes.tablePageSpace, `${height}`);
  row.setAttribute("aria-hidden", "true");
  row.setAttribute("contenteditable", "false");
  row.style.height = `${height}px`;

  const cell = document.createElement("td");
  cell.colSpan = columns;
  cell.style.height = `${height}px`;
  row.append(cell);
  return row;
}

function repeatedHeader(view: EditorView, rowPos: number): HTMLElement {
  const source = view.nodeDOM(rowPos);
  const row =
    source instanceof HTMLElement && source.tagName === "TR"
      ? (source.cloneNode(true) as HTMLElement)
      : document.createElement("tr");
  row.setAttribute(editorAttributes.tableRepeatedHeader, "");
  row.setAttribute("aria-hidden", "true");
  row.setAttribute("contenteditable", "false");
  row.removeAttribute("id");
  row.querySelectorAll("[id]").forEach((element) => {
    element.removeAttribute("id");
  });
  return row;
}

/**
 * The spacer and the repeated headers a table's cuts are drawn as.
 *
 * What each one needs beyond the cut itself - which rows repeat, how wide a spacer has to be - is
 * read back off the table as it stands now, so an edit to a header row redraws the projection of
 * it without the layout having to be run again.
 */
function tableContinuations(
  table: PMNode,
  tablePos: number,
  cuts: readonly PageCut[],
  into: Decoration[]
): void {
  if (cuts.length === 0 || table.type.spec.tableRole !== "table") return;
  const headerRows = headerRowsOf(table, tablePos);
  const columns = columnCount(table);
  for (const cut of cuts) {
    into.push(
      Decoration.widget(cut.at, () => tableSpace(cut.height, columns), {
        key: `table-page-space-${cut.at}-${cut.height}-${columns}`,
        side: -100,
      })
    );
    headerRows.forEach((rowPos, headerIndex) => {
      into.push(
        // No `key`: a keyed widget is held to be the same one and left alone, and this one is a
        // copy of a row that the very edit rebuilding these decorations may have just changed
        Decoration.widget(cut.at, (view) => repeatedHeader(view, rowPos), {
          side: -90 + headerIndex,
        })
      );
    });
  }
}

/**
 * Every page break carries a space, an empty one until a measurement says otherwise: the
 * measurement reads where a break stands off that very element, so it has to be there before the
 * first one is taken.
 */
function decorationsFor(
  doc: PMNode,
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
    const blockCuts = byBlock.get(offset) ?? [];
    breakSpaces(node, offset, blockCuts, decorations);
    tableContinuations(node, offset, blockCuts, decorations);
  });
  return DecorationSet.create(doc, decorations);
}

function marksFor(
  doc: PMNode,
  pushes: readonly BlockPush[],
  cuts: readonly PageCut[]
): PageMarks {
  return { pushes, cuts, decorations: decorationsFor(doc, pushes, cuts) };
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
        other.marginTop === push.marginTop
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

export function pageDecorations(): Plugin<PageMarks> {
  return new Plugin<PageMarks>({
    key: marksKey,
    state: {
      init: (_config, state) => marksFor(state.doc, [], []),
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
          value.pushes.map((push) => ({
            ...push,
            pos: tr.mapping.map(push.pos),
          })),
          value.cuts.flatMap((cut) => {
            const mapped = tr.mapping.mapResult(cut.at, 1);
            const node = tr.doc.nodeAt(mapped.pos);
            const kept = isPageBreakNode(node) || isRow(node);
            return mapped.deleted || !kept ? [] : [{ ...cut, at: mapped.pos }];
          })
        );
      },
    },
    props: {
      decorations: (state) => marksKey.getState(state)?.decorations,
    },
  });
}

function current(view: EditorView): PageMarks {
  return marksKey.getState(view.state) ?? marksFor(view.state.doc, [], []);
}

/** One transaction, or none when the same marks are already applied */
export function setPageMarks(view: EditorView, next: PageMarksInput): void {
  if (samePageMarks(current(view), next)) return;
  view.dispatch(
    view.state.tr
      .setMeta(marksKey, marksFor(view.state.doc, next.pushes, next.cuts))
      .setMeta("addToHistory", false)
  );
}
