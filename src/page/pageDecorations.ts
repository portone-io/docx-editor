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
import type { BlockPush, BreakSpace, TableContinuation } from "./pageLayout";

/**
 * One break's space while the document changes under it.
 *
 * The layout counts breaks inside the block it measured (`page/pageLayout`), and an edit anywhere
 * before a break shifts that count, so what is kept here is the position of the `br` the space
 * belongs to: a position maps through an edit, an ordinal does not.
 */
interface SpaceMark {
  /** Where the `br` this space is put on stands */
  at: number;
  height: number;
}

interface PageMarks {
  pushes: readonly BlockPush[];
  spaces: readonly SpaceMark[];
  tableContinuations: readonly TableContinuation[];
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
 */
function pageBreaksIn(block: PMNode, blockPos: number): BreakAt[] {
  const found: BreakAt[] = [];
  if (block.type !== docxSchema.nodes.paragraph) return found;
  block.forEach((child, offset) => {
    if (isPageBreakNode(child)) {
      found.push({ at: blockPos + 1 + offset, size: child.nodeSize });
    }
  });
  return found;
}

function breakSpaces(
  block: PMNode,
  blockPos: number,
  heights: ReadonlyMap<number, number>,
  into: Decoration[]
): void {
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

/**
 * The layout's spaces as the marks that follow their breaks: each `(block, ordinal)` resolved
 * against the document it was measured on. One naming a break the document no longer holds is
 * dropped.
 */
function spaceMarks(doc: PMNode, spaces: readonly BreakSpace[]): SpaceMark[] {
  return spaces.flatMap((space) => {
    const block = doc.nodeAt(space.pos);
    const found = block ? pageBreaksIn(block, space.pos)[space.index] : null;
    return found ? [{ at: found.at, height: space.height }] : [];
  });
}

function tableSpace(continuation: TableContinuation): HTMLElement {
  const row = document.createElement("tr");
  row.setAttribute(editorAttributes.tablePageSpace, `${continuation.height}`);
  row.setAttribute("aria-hidden", "true");
  row.setAttribute("contenteditable", "false");
  row.style.height = `${continuation.height}px`;

  const cell = document.createElement("td");
  cell.colSpan = continuation.columns;
  cell.style.height = `${continuation.height}px`;
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

function tableContinuationDecorations(
  continuations: readonly TableContinuation[],
  into: Decoration[]
): void {
  continuations.forEach((continuation) => {
    into.push(
      Decoration.widget(continuation.pos, () => tableSpace(continuation), {
        key: `table-page-space-${continuation.pos}-${continuation.height}-${continuation.columns}`,
        side: -100,
      })
    );
    continuation.headerRows.forEach((rowPos, headerIndex) => {
      into.push(
        Decoration.widget(
          continuation.pos,
          (view) => repeatedHeader(view, rowPos),
          {
            key:
              `table-repeated-header-${continuation.pos}-${headerIndex}-` +
              continuation.headerSignature,
            side: -90 + headerIndex,
          }
        )
      );
    });
  });
}

/**
 * Every page break carries a space, an empty one until a measurement says otherwise: the
 * measurement reads where a break stands off that very element, so it has to be there before the
 * first one is taken.
 */
function decorationsFor(
  doc: PMNode,
  pushes: readonly BlockPush[],
  spaces: readonly SpaceMark[],
  tableContinuations: readonly TableContinuation[]
): DecorationSet {
  const byPos = new Map(pushes.map((push) => [push.pos, push]));
  const heights = new Map(spaces.map((space) => [space.at, space.height]));
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
    breakSpaces(node, offset, heights, decorations);
  });
  tableContinuationDecorations(tableContinuations, decorations);
  return DecorationSet.create(doc, decorations);
}

function marksFor(
  doc: PMNode,
  pushes: readonly BlockPush[],
  spaces: readonly SpaceMark[],
  tableContinuations: readonly TableContinuation[]
): PageMarks {
  return {
    pushes,
    spaces,
    tableContinuations,
    decorations: decorationsFor(doc, pushes, spaces, tableContinuations),
  };
}

function samePushes(a: readonly BlockPush[], b: readonly BlockPush[]): boolean {
  return (
    a.length === b.length &&
    a.every((push, index) => {
      const other = b[index];
      return (
        other !== undefined &&
        other.pos === push.pos &&
        other.marginTop === push.marginTop
      );
    })
  );
}

function sameSpaces(a: readonly SpaceMark[], b: readonly SpaceMark[]): boolean {
  return (
    a.length === b.length &&
    a.every((space, index) => {
      const other = b[index];
      return (
        other !== undefined &&
        other.at === space.at &&
        other.height === space.height
      );
    })
  );
}

function sameTableContinuations(
  a: readonly TableContinuation[],
  b: readonly TableContinuation[]
): boolean {
  return (
    a.length === b.length &&
    a.every((continuation, index) => {
      const other = b[index];
      return (
        other !== undefined &&
        other.pos === continuation.pos &&
        other.height === continuation.height &&
        other.columns === continuation.columns &&
        other.headerSignature === continuation.headerSignature &&
        other.headerRows.length === continuation.headerRows.length &&
        other.headerRows.every(
          (rowPos, rowIndex) => rowPos === continuation.headerRows[rowIndex]
        )
      );
    })
  );
}

export function pageDecorations(): Plugin<PageMarks> {
  return new Plugin<PageMarks>({
    key: marksKey,
    state: {
      init: (_config, state) => marksFor(state.doc, [], [], []),
      apply(tr, value) {
        const next = tr.getMeta(marksKey);
        if (next) return next;
        if (!tr.docChanged) return value;
        // Positions shift when the text changes, and a break the edit has just put in has no
        // space yet, so the marks are laid out again over the new document. A space whose `br`
        // the edit took away goes with it, rather than landing on whatever the mapping now points
        // at
        return marksFor(
          tr.doc,
          value.pushes.map((push) => ({
            ...push,
            pos: tr.mapping.map(push.pos),
          })),
          value.spaces.flatMap((space) => {
            const at = tr.mapping.map(space.at);
            return isPageBreakNode(tr.doc.nodeAt(at)) ? [{ ...space, at }] : [];
          }),
          value.tableContinuations.flatMap((continuation) => {
            const pos = tr.mapping.map(continuation.pos);
            if (tr.doc.nodeAt(pos)?.type.spec.tableRole !== "row") return [];
            const headerRows = continuation.headerRows
              .map((rowPos) => tr.mapping.map(rowPos))
              .filter(
                (rowPos) => tr.doc.nodeAt(rowPos)?.type.spec.tableRole === "row"
              );
            return [{ ...continuation, pos, headerRows }];
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
  return marksKey.getState(view.state) ?? marksFor(view.state.doc, [], [], []);
}

function dispatch(view: EditorView, next: PageMarks): void {
  view.dispatch(
    view.state.tr.setMeta(marksKey, next).setMeta("addToHistory", false)
  );
}

/** Does nothing when the same values are already applied */
export function setPagePushes(
  view: EditorView,
  pushes: readonly BlockPush[]
): void {
  const marks = current(view);
  if (samePushes(marks.pushes, pushes)) return;
  dispatch(
    view,
    marksFor(view.state.doc, pushes, marks.spaces, marks.tableContinuations)
  );
}

/** Does nothing when the same values are already applied */
export function setPageBreakSpaces(
  view: EditorView,
  spaces: readonly BreakSpace[]
): void {
  const marks = current(view);
  const next = spaceMarks(view.state.doc, spaces);
  if (sameSpaces(marks.spaces, next)) return;
  dispatch(
    view,
    marksFor(view.state.doc, marks.pushes, next, marks.tableContinuations)
  );
}

/** Does nothing when the same table continuations are already applied */
export function setTableContinuations(
  view: EditorView,
  continuations: readonly TableContinuation[]
): void {
  const marks = current(view);
  if (sameTableContinuations(marks.tableContinuations, continuations)) return;
  dispatch(
    view,
    marksFor(view.state.doc, marks.pushes, marks.spaces, continuations)
  );
}
