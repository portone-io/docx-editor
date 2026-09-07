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
 *
 * The one walk answers two callers: `withUniqueIdentities` hands the writer the settled document
 * and throws over a block that cannot yield, and `identityProblems` runs the same walk for
 * `docx/invariants` and lists those blocks instead, so what the query reports and what the write
 * refuses cannot come apart.
 */

import { Fragment, type Mark, type Node as PMNode } from "prosemirror-model";
import { DocxExportError, type DocxExportErrorCode } from "../ooxml/errors";
import { qualify } from "../ooxml/names";
import { parseAttrs } from "../ooxml/tagScan";
import { docxSchema } from "../schema";
import { withoutParagraphIds } from "./cloning";
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

/**
 * The block a node was opened from, which `docx/exportDocx` writes the original bytes back for.
 * A second node claiming the same block is written from its own attrs instead, and a block that
 * exists only as its original XML has nothing to be written from, so it cannot yield.
 */
export const sourceBlockRule: IdentityRule = {
  name: "sourceBlock",
  visit(block, written) {
    const srcId: unknown = block.attrs.srcId;
    if (typeof srcId !== "string" && typeof srcId !== "number") return block;
    const name = String(srcId);
    if (!written.has(name)) {
      written.add(name);
      return block;
    }
    if (block.type.isInGroup("preserved")) return null;
    return block.type.create(
      { ...block.attrs, srcId: null },
      block.content,
      block.marks
    );
  },
};

const PARA_ID = qualify("w14", "paraId");

/** `ST_LongHexNumber` (§17.18.50): eight hexadecimal digits, spelled in either case */
const LONG_HEX = /^[0-9A-Fa-f]{8}$/;

/**
 * The paragraph identifier this paragraph carries, read as the number it spells so that two
 * spellings of one value are one name. null where there is none, and also where the value is
 * not one [MS-DOCX] §2.6.2.3 allows. This pass leaves invalid input as it came rather than
 * attempting to repair it.
 */
function paragraphIdOf(pAttrs: unknown): number | null {
  if (typeof pAttrs !== "string") return null;
  const value = parseAttrs(pAttrs)?.find(([name]) => name === PARA_ID)?.[1];
  if (value === undefined || !LONG_HEX.test(value)) return null;
  const id = Number.parseInt(value, 16);
  return id > 0 && id < 0x80000000 ? id : null;
}

/**
 * A later paragraph carrying an identifier already written goes out without one, and without its
 * `w14:textId` too, which may not stand without a `w14:paraId` beside it. The walker reaches the
 * paragraphs inside a table, so the rule reads the one block it is handed.
 */
export const paragraphIdRule: IdentityRule = {
  name: "paragraphId",
  visit(block, written) {
    if (block.type !== docxSchema.nodes.paragraph) return block;
    const id = paragraphIdOf(block.attrs.pAttrs);
    if (id === null) return block;
    const name = String(id);
    if (!written.has(name)) {
      written.add(name);
      return block;
    }
    return block.type.create(
      { ...block.attrs, pAttrs: withoutParagraphIds(block.attrs.pAttrs) },
      block.content,
      block.marks
    );
  },
};

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

/**
 * In the order they apply to one block: the source rule first, so that a second claimant of a
 * block is already a rebuilt paragraph by the time the paragraph rule takes its identifier away.
 */
export const IDENTITY_RULES: readonly IdentityRule[] = [
  sourceBlockRule,
  paragraphIdRule,
  controlRule,
];

/** A block that cannot yield: what the export refuses it with, and where it stands in the document */
export interface IdentityProblem {
  readonly code: DocxExportErrorCode;
  readonly message: string;
  readonly pos: number;
}

function refusalOf(block: PMNode, pos: number): IdentityProblem {
  return {
    code: "unsupported-content",
    message: `a preserved block stands in two places (${block.type.name})`,
    pos,
  };
}

/** One pass over the blocks: the rules, the names each has written so far, and what to do with a block that cannot yield */
interface Walk {
  readonly rules: readonly IdentityRule[];
  readonly written: readonly Set<string>[];
  readonly refuse: (problem: IdentityProblem) => void;
}

/**
 * The block with every rule applied in turn, each reading what the one before it made of it. A
 * block that cannot yield is handed to `refuse` and stands as it is for the rules after that one.
 */
function visitBlock(block: PMNode, pos: number, walk: Walk): PMNode {
  let current = block;
  for (const [index, rule] of walk.rules.entries()) {
    const next = rule.visit(current, walk.written[index]);
    if (next === null) {
      walk.refuse(refusalOf(block, pos));
      return current;
    }
    current = next;
  }
  return current;
}

function rewriteBlock(block: PMNode, pos: number, walk: Walk): PMNode {
  const visited = visitBlock(block, pos, walk);
  if (visited.inlineContent || visited.childCount === 0) return visited;
  const children: PMNode[] = [];
  let changed = false;
  visited.forEach((child, offset) => {
    const next = rewriteBlock(child, pos + 1 + offset, walk);
    changed ||= next !== child;
    children.push(next);
  });
  return changed ? visited.copy(Fragment.fromArray(children)) : visited;
}

/** The blocks in document order, one `written` set per rule, and the document itself where nothing changed */
function walkBlocks(
  doc: PMNode,
  rules: readonly IdentityRule[],
  refuse: Walk["refuse"]
): PMNode {
  const walk: Walk = {
    rules,
    written: rules.map(() => new Set<string>()),
    refuse,
  };
  const blocks: PMNode[] = [];
  let changed = false;
  doc.forEach((block, offset) => {
    const next = rewriteBlock(block, offset, walk);
    changed ||= next !== block;
    blocks.push(next);
  });
  return changed ? doc.copy(Fragment.fromArray(blocks)) : doc;
}

/** The document with every later claimant of a name released, in document order */
export function withUniqueIdentities(
  doc: PMNode,
  rules: readonly IdentityRule[] = IDENTITY_RULES
): PMNode {
  return walkBlocks(doc, rules, (problem) => {
    throw new DocxExportError(problem.code, problem.message);
  });
}

/**
 * Every block `withUniqueIdentities` would refuse over, in document order, each where it stands;
 * empty when the pass would go through. The pass is run as the export runs it and its settled
 * document is dropped, so a caller asking ahead of the write reads the same refusal it would throw.
 */
export function identityProblems(
  doc: PMNode,
  rules: readonly IdentityRule[] = IDENTITY_RULES
): readonly IdentityProblem[] {
  const problems: IdentityProblem[] = [];
  walkBlocks(doc, rules, (problem) => problems.push(problem));
  return problems;
}
