/**
 * Keeping one content control from going out twice under the same name.
 *
 * A control cannot cross a paragraph, so splitting a paragraph in the middle of one leaves that
 * same control standing in two, and dropping unmarked text into the middle of one breaks it in
 * two within the paragraph. Each piece then writes a `w:sdt` of its own (see `docx/sdt` for what
 * a copy must not carry along).
 *
 * A document nobody broke apart comes back node for node as it was, which is what keeps an
 * unedited block going out as its original XML.
 */

import { Fragment, type Mark, type Node as PMNode } from "prosemirror-model";
import { docxSchema } from "../schema";
import { copiedControlPrefix, newControlId } from "./sdt";

function sdtMarkOf(node: PMNode): Mark | null {
  return node.marks.find((mark) => mark.type === docxSchema.marks.sdt) ?? null;
}

/** What tells one control apart from another, wherever in the document it turns up */
function controlName(mark: Mark): string {
  const key: unknown = mark.attrs.sdtKey;
  const prefix: unknown = mark.attrs.sdtPrefix;
  return `${typeof key === "number" ? key : 0} ${typeof prefix === "string" ? prefix : ""}`;
}

function copiedMark(mark: Mark): Mark {
  const prefix: unknown = mark.attrs.sdtPrefix;
  if (typeof prefix !== "string") return mark;
  return mark.type.create({
    ...mark.attrs,
    sdtPrefix: copiedControlPrefix(prefix, newControlId()),
  });
}

/** One stretch of inline nodes wearing the same control mark */
interface Control {
  mark: Mark;
  /** The mark to wear in its place. null while this is the first time the control goes out */
  copy: Mark | null;
}

function claim(mark: Mark, written: Set<string>): Control {
  const name = controlName(mark);
  const copy = written.has(name) ? copiedMark(mark) : null;
  written.add(name);
  return { mark, copy };
}

function rewriteParagraph(paragraph: PMNode, written: Set<string>): PMNode {
  const inline: PMNode[] = [];
  let running: Control | null = null;
  let renamed = false;

  paragraph.forEach((child) => {
    const mark = sdtMarkOf(child);
    if (!mark) {
      running = null;
      inline.push(child);
      return;
    }
    const control = running?.mark.eq(mark) ? running : claim(mark, written);
    running = control;
    if (!control.copy) {
      inline.push(child);
      return;
    }
    renamed = true;
    inline.push(
      child.mark(control.copy.addToSet(mark.removeFromSet(child.marks)))
    );
  });

  return renamed ? paragraph.copy(Fragment.fromArray(inline)) : paragraph;
}

function rewriteBlock(node: PMNode, written: Set<string>): PMNode {
  if (node.type.name === "paragraph") return rewriteParagraph(node, written);
  if (node.childCount === 0) return node;
  const children = node.children.map((child) => rewriteBlock(child, written));
  return children.every((child, index) => child === node.child(index))
    ? node
    : node.copy(Fragment.fromArray(children));
}

/** The document with every copy of a control opening under a name of its own */
export function withUniqueControls(doc: PMNode): PMNode {
  return rewriteBlock(doc, new Set());
}
