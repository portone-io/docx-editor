/**
 * Settling, just before export, every name that may be held by one node only.
 *
 * An edit can leave one name standing in two places: splitting a paragraph inside a content
 * control leaves that control in both halves, and dropping a copy of a paragraph carries the
 * original's identifiers along with it. Each name a document must keep unique is one rule here,
 * and the pass walks the blocks once in document order, so the first node to claim a name keeps it
 * and every later claimant gives it up.
 *
 * A document nobody broke apart comes back node for node as it was, which is what keeps an
 * unedited block going out as its original XML.
 */

import { Fragment, type Mark, type Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import { copiedControlPrefix, newControlId } from "./sdt";

export interface IdentityRule {
  readonly name: string;
  /**
   * The block to write in place of `block`, with every name it claims that `written` already holds
   * given up, and the names it keeps added to `written`. Returns `block` itself when nothing
   * changes. null means the block cannot yield and the export must refuse. A block's children are
   * not visited here; the walker visits them after this call.
   */
  visit(block: PMNode, written: Set<string>): PMNode | null;
}

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

/**
 * The several runs of one control are that one control, so the paragraph's inline nodes are
 * walked here rather than claimed one by one: a control is a stretch of neighbours wearing the
 * same mark, and it is broken in two only where an unmarked node stands between them.
 */
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

/**
 * A content control cannot cross a paragraph, so splitting a paragraph in the middle of one leaves
 * that same control standing in two, and dropping unmarked text into the middle of one breaks it
 * in two within the paragraph. Each piece after the first opens as a copy with a `w:id` of its
 * own (see `docx/sdt` for what a copy must not carry along).
 */
export const controlRule: IdentityRule = {
  name: "control",
  visit(block, written) {
    return block.type === docxSchema.nodes.paragraph
      ? rewriteParagraph(block, written)
      : block;
  },
};

export const IDENTITY_RULES: readonly IdentityRule[] = [controlRule];

/** The block with every rule applied in turn, each reading what the one before it made of it */
function visitBlock(
  block: PMNode,
  rules: readonly IdentityRule[],
  written: readonly Set<string>[]
): PMNode {
  return rules.reduce((current, rule, index) => {
    const next = rule.visit(current, written[index]);
    if (next === null) {
      throw new DocxExportError(
        "unsupported-content",
        `a preserved block stands in two places (${block.type.name})`
      );
    }
    return next;
  }, block);
}

function rewriteBlock(
  block: PMNode,
  rules: readonly IdentityRule[],
  written: readonly Set<string>[]
): PMNode {
  const visited = visitBlock(block, rules, written);
  if (visited.inlineContent || visited.childCount === 0) return visited;
  const children = visited.children.map((child) =>
    rewriteBlock(child, rules, written)
  );
  return children.every((child, index) => child === visited.child(index))
    ? visited
    : visited.copy(Fragment.fromArray(children));
}

/** The document with every later claimant of a name released, in document order, one `written` set per rule */
export function withUniqueIdentities(
  doc: PMNode,
  rules: readonly IdentityRule[] = IDENTITY_RULES
): PMNode {
  const written = rules.map(() => new Set<string>());
  const blocks = doc.children.map((block) =>
    rewriteBlock(block, rules, written)
  );
  return blocks.every((block, index) => block === doc.child(index))
    ? doc
    : doc.copy(Fragment.fromArray(blocks));
}
