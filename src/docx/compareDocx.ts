/**
 * What changed between two files, block by block: the answer `onlyCommentsChangedBy` withholds
 * when all it says is whether anything but comments did.
 *
 * Two files opened side by side share no block key, and a wrapper's key counts through the
 * document it was opened from, so the model's own `sameSource` calls every block of the second
 * file changed. Blocks are compared as the writer puts them out instead (`./storyProjection`),
 * which is the comparison the returned-file verifier holds a story to, and the package around
 * them is read through the verifier's own pieces (`./protectionPolicy`).
 */

import type { Node as PMNode } from "prosemirror-model";
import { withXmlParser, type XmlParser } from "../ooxml/xml";
import { SDT_BLOCK_NODE, SDT_EMPTY_BLOCK_NODE } from "../schema/controlAttrs";
import { withoutComments } from "../schema/protection";
import { commentReferencesIn } from "./comments/model";
import { commentsPolicy } from "./comments/policy";
import { type DocxBytes, importDocx } from "./importDocx";
import { sameBytes } from "./packageParts";
import {
  aroundTheBlocks,
  declarationsKept,
  policyPartPaths,
} from "./protectionPolicy";
import { type Alignment, alignSequences } from "./sequenceDiff";
import { storyKey, storyOf, storyText } from "./story";
import { comparableBlock, type Story } from "./storyProjection";
import { diffText, type TextEdit } from "./textDiff";

/** What stands at the top level of a body: `control` is a block-level content control, `preserved` a block kept as the XML it arrived as */
export type BlockKind = "paragraph" | "table" | "control" | "preserved";

/** A row of a changed table: an added row by its index in the revised table, a removed one by its index in the original */
export type RowChange =
  | { kind: "added"; index: number; cells: string[] }
  | { kind: "removed"; index: number; cells: string[] }
  | {
      kind: "changed";
      original: { index: number; cells: string[] };
      revised: { index: number; cells: string[] };
    };

/**
 * A top-level block of the body that differs. An added block is placed by its `pos` in the revised
 * document, a removed one by its `pos` in the original. A changed block is `formatting` where its
 * text reads the same on both sides, and `text` otherwise, with the `edits` that turn one text
 * into the other. A changed table also lists its `rows` whose cells read differently.
 */
export type BlockChange =
  | { kind: "added"; block: BlockKind; pos: number; text: string }
  | { kind: "removed"; block: BlockKind; pos: number; text: string }
  | {
      kind: "changed";
      block: Exclude<BlockKind, "table">;
      original: { pos: number; text: string };
      revised: { pos: number; text: string };
      change: "text";
      edits: TextEdit[];
    }
  | {
      kind: "changed";
      block: Exclude<BlockKind, "table">;
      original: { pos: number; text: string };
      revised: { pos: number; text: string };
      change: "formatting";
    }
  | {
      kind: "changed";
      block: "table";
      original: { pos: number; text: string };
      revised: { pos: number; text: string };
      change: "text";
      edits: TextEdit[];
      rows: RowChange[];
    }
  | {
      kind: "changed";
      block: "table";
      original: { pos: number; text: string };
      revised: { pos: number; text: string };
      change: "formatting";
      rows: RowChange[];
    };

/**
 * A comment or reply that differs. Comments are matched by id, and one whose author differs
 * between the two files is a different comment under a reused id, so it is removed and added.
 */
export type CommentChange =
  | {
      kind: "added";
      id: string;
      author: string | null;
      authorId: string | null;
      text: string;
    }
  | {
      kind: "removed";
      id: string;
      author: string | null;
      authorId: string | null;
      text: string;
    }
  | {
      kind: "changed";
      id: string;
      author: string | null;
      authorId: string | null;
      original: string;
      revised: string;
    };

/** A package part beside the body and the comment parts whose bytes differ, or the main part where it differs outside the body blocks */
export interface PartChange {
  part: string;
  kind: "added" | "removed" | "changed";
}

/** What `compareDocx` found: the body's blocks in document order, the comments, and the parts sorted by path */
export interface DocxComparison {
  blocks: BlockChange[];
  comments: CommentChange[];
  parts: PartChange[];
}

/** Every top-level node type the schema holds; one it gains fails the test that walks the block group */
export const BLOCK_KINDS: Readonly<Record<string, BlockKind>> = {
  paragraph: "paragraph",
  table: "table",
  [SDT_BLOCK_NODE]: "control",
  [SDT_EMPTY_BLOCK_NODE]: "control",
  rawBlock: "preserved",
};

function blockKind(node: PMNode): BlockKind {
  const kind = BLOCK_KINDS[node.type.name];
  if (kind === undefined) {
    throw new Error(`no block kind for ${node.type.name}`);
  }
  return kind;
}

interface Block {
  node: PMNode;
  pos: number;
  block: BlockKind;
  text: string;
  identity: string;
}

/**
 * A block the writer cannot put out has nothing to be compared by but its text, so it takes an
 * identity no other block of either side can share and is never kept as it was.
 */
function blocksOf(story: Story, side: "original" | "revised"): Block[] {
  const blocks: Block[] = [];
  story.doc.forEach((node, pos, index) => {
    blocks.push({
      node,
      pos,
      block: blockKind(node),
      text: storyText(node),
      identity:
        comparableBlock(node, story.session, withoutComments) ??
        `\u0000${side}:${index}`,
    });
  });
  return blocks;
}

interface Row {
  index: number;
  cells: string[];
  identity: string;
  text: string;
}

type Difference = Exclude<Alignment, { kind: "kept" }>;

function isDifference(step: Alignment): step is Difference {
  return step.kind !== "kept";
}

function rowAt({ index, cells }: Row): { index: number; cells: string[] } {
  return { index, cells };
}

function rowsOf(table: PMNode): Row[] {
  return table.children.map((row, index) => {
    const cells = row.children.map(storyText);
    const text = JSON.stringify(cells);
    return { index, cells, identity: text, text };
  });
}

function rowChanges(original: PMNode, revised: PMNode): RowChange[] {
  const was = rowsOf(original);
  const now = rowsOf(revised);
  const rowChange = (step: Difference): RowChange => {
    switch (step.kind) {
      case "added":
        return { kind: "added", ...rowAt(now[step.revised]) };
      case "removed":
        return { kind: "removed", ...rowAt(was[step.original]) };
      case "changed":
        return {
          kind: "changed",
          original: rowAt(was[step.original]),
          revised: rowAt(now[step.revised]),
        };
    }
  };
  return alignSequences(was, now).filter(isDifference).map(rowChange);
}

function blockAt({ pos, text }: Block): { pos: number; text: string } {
  return { pos, text };
}

function changedBlock(
  from: Block,
  to: Block,
  change: "text" | "formatting"
): BlockChange {
  const original = blockAt(from);
  const revised = blockAt(to);
  if (to.block === "table") {
    const rows = rowChanges(from.node, to.node);
    return change === "text"
      ? {
          kind: "changed",
          block: to.block,
          original,
          revised,
          change,
          edits: diffText(from.text, to.text),
          rows,
        }
      : { kind: "changed", block: to.block, original, revised, change, rows };
  }
  return change === "text"
    ? {
        kind: "changed",
        block: to.block,
        original,
        revised,
        change,
        edits: diffText(from.text, to.text),
      }
    : { kind: "changed", block: to.block, original, revised, change };
}

function blockChanges(original: Story, revised: Story): BlockChange[] {
  const was = blocksOf(original, "original");
  const now = blocksOf(revised, "revised");
  const blockChange = (step: Difference): BlockChange => {
    switch (step.kind) {
      case "added": {
        const block = now[step.revised];
        return { kind: "added", block: block.block, ...blockAt(block) };
      }
      case "removed": {
        const block = was[step.original];
        return { kind: "removed", block: block.block, ...blockAt(block) };
      }
      case "changed":
        return changedBlock(was[step.original], now[step.revised], step.change);
    }
  };
  return alignSequences(was, now, (block) => block.block)
    .filter(isDifference)
    .map(blockChange);
}

interface CommentRead {
  id: string;
  author: string | null;
  authorId: string | null;
  text: string;
}

/** Every comment and reply the body refers to, in document order, each thread's replies after it */
function commentsOf(doc: PMNode): Map<string, CommentRead> {
  const read = new Map<string, CommentRead>();
  const add = (comment: Omit<CommentRead, "text">) => {
    if (read.has(comment.id)) return;
    read.set(comment.id, {
      id: comment.id,
      author: comment.author,
      authorId: comment.authorId,
      text: storyText(storyOf(doc, storyKey("comment", comment.id))),
    });
  };
  for (const comment of commentReferencesIn(doc).values()) {
    add(comment);
    comment.replies.forEach(add);
  }
  return read;
}

/** Whether two comments under one id were written by the same person, as far as the files say */
function sameAuthor(original: CommentRead, revised: CommentRead): boolean {
  return original.authorId === null && revised.authorId === null
    ? original.author === revised.author
    : original.authorId === revised.authorId;
}

function commentChanges(original: PMNode, revised: PMNode): CommentChange[] {
  const was = commentsOf(original);
  const now = commentsOf(revised);
  const changes: CommentChange[] = [];
  for (const [id, comment] of was) {
    const current = now.get(id);
    if (current === undefined || !sameAuthor(comment, current)) {
      changes.push({ kind: "removed", ...comment });
    } else if (current.text !== comment.text) {
      changes.push({
        kind: "changed",
        id,
        author: current.author,
        authorId: current.authorId,
        original: comment.text,
        revised: current.text,
      });
    }
  }
  for (const [id, comment] of now) {
    const earlier = was.get(id);
    if (earlier === undefined || !sameAuthor(earlier, comment)) {
      changes.push({ kind: "added", ...comment });
    }
  }
  return changes;
}

function partChanges(original: Story, revised: Story): PartChange[] {
  const was = original.session;
  const now = revised.session;
  const sameMainPart = was.mainPartPath === now.mainPartPath;
  const compared = new Set([
    was.mainPartPath,
    now.mainPartPath,
    ...policyPartPaths(commentsPolicy, was),
    ...policyPartPaths(commentsPolicy, now),
  ]);
  const declarations = sameMainPart
    ? declarationsKept(commentsPolicy, was, now, "gained or dropped")
    : new Map<string, boolean>();
  const changes: PartChange[] = sameMainPart
    ? []
    : [
        { part: was.mainPartPath, kind: "removed" },
        { part: now.mainPartPath, kind: "added" },
      ];
  for (const part of new Set([...was.parts.keys(), ...now.parts.keys()])) {
    if (compared.has(part) || declarations.get(part) === true) continue;
    const arrived = was.parts.get(part);
    const current = now.parts.get(part);
    if (
      arrived !== undefined &&
      current !== undefined &&
      sameBytes(arrived, current)
    ) {
      continue;
    }
    changes.push({
      part,
      kind:
        arrived === undefined
          ? "added"
          : current === undefined
            ? "removed"
            : "changed",
    });
  }
  if (sameMainPart && aroundTheBlocks(original) !== aroundTheBlocks(revised)) {
    changes.push({ part: was.mainPartPath, kind: "changed" });
  }
  return changes.sort((a, b) =>
    a.part < b.part ? -1 : a.part > b.part ? 1 : 0
  );
}

/**
 * Lists what differs between two DOCX files: the top-level blocks of the body, the comments and
 * replies, and the package parts beside them. A block is compared with its comment markers taken
 * out, so a comment added or removed is reported among the comments alone.
 *
 * @param original - the file to compare against, as an `ArrayBuffer` or `Uint8Array`
 * @param revised - the file to compare, as an `ArrayBuffer` or `Uint8Array`
 * @param options - `xmlParser` for a runtime with no `DOMParser` global
 * @throws `DocxImportError` when either file is not a readable DOCX
 */
export function compareDocx(
  original: DocxBytes,
  revised: DocxBytes,
  { xmlParser }: { xmlParser?: XmlParser } = {}
): DocxComparison {
  return withXmlParser(xmlParser, () => {
    const was = importDocx(original);
    const now = importDocx(revised);
    return {
      blocks: blockChanges(was, now),
      comments: commentChanges(was.doc, now.doc),
      parts: partChanges(was, now),
    };
  });
}
