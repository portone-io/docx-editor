/**
 * Renders a table from the schema's `toDOM`, and keeps ProseMirror from reverting the `colgroup`
 * we edit ourselves while column widths are being dragged.
 *
 * When an attribute changes inside the editor DOM, ProseMirror treats it as "DOM out of sync
 * with the document" and redraws that spot from the document. That is why the `col` widths
 * written by the preview get wiped on the next microtask, and dragging moves nothing on screen.
 * `prosemirror-tables`' `TableView` does the same thing for the same reason (MIT; we borrowed
 * only the approach).
 */

import { DOMSerializer, type Node as PMNode } from "prosemirror-model";
import type { NodeView, ViewMutationRecord } from "prosemirror-view";

export class TableNodeView implements NodeView {
  readonly dom: HTMLElement;
  readonly contentDOM: HTMLElement | undefined;
  /** A table whose grid is unknown has no colgroup */
  private readonly colgroup: Element | null;

  constructor(node: PMNode) {
    const toDOM = node.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("the schema has no way to render a table");
    }
    const rendered = DOMSerializer.renderSpec(document, toDOM(node));
    if (!(rendered.dom instanceof HTMLElement)) {
      throw new Error("the table was not rendered as an element");
    }
    this.dom = rendered.dom;
    this.contentDOM = rendered.contentDOM;
    this.colgroup = this.dom.querySelector(":scope > colgroup");
  }

  /**
   * Attribute changes on the table and its grid are ones we made ourselves, so they are not
   * re-read from the document. Changes that happen inside cells are seen by ProseMirror as usual.
   */
  ignoreMutation(record: ViewMutationRecord): boolean {
    if (record.type !== "attributes") return false;
    return (
      record.target === this.dom ||
      this.colgroup?.contains(record.target) === true
    );
  }
}
