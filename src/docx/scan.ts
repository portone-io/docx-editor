/**
 * Finds the span each individual block occupies in the raw text of a part and slices it out.
 *
 * A block that was not edited is exported by writing this fragment back out untouched,
 * which keeps its original bytes intact. The body of the main part is one such run of blocks;
 * a comment and a footnote hold another (`docx/story`), and all of them are sliced here so that
 * one story cannot end up compared or written by rules another does not follow.
 */

import { DocxImportError } from "../ooxml/errors";
import { readTag, type Tag } from "../ooxml/tagScan";
import { localPart } from "../ooxml/xml";

export interface BlockSlice {
  /** The tag name exactly as written (e.g. "w:p", "w:tbl", "w:sectPr") */
  name: string;
  /** The original XML fragment. It also carries the whitespace that sat between this block and the one before it */
  xml: string;
}

export interface BlockScan {
  /** Everything in the raw text before the first block, the container's own opening tag included */
  prefix: string;
  blocks: BlockSlice[];
  /** Everything in the raw text after the last block, the container's closing tag included */
  suffix: string;
}

/** The tag whose children are the blocks, asked of every opening tag with the depth it stands at */
export type ContainerTest = (tag: Tag, depth: number) => boolean;

/** The tag at this `<`, refusing a document whose text cannot be read tag by tag */
function tagAt(source: string, lt: number): Tag {
  const tag = readTag(source, lt);
  if (!tag) {
    throw new DocxImportError(
      "malformed-xml",
      "a tag cannot be read to its end"
    );
  }
  return tag;
}

/** The main part's body, which is the container the document's own blocks stand in */
function isBodyTag(tag: Tag, depth: number): boolean {
  return localPart(tag.name) === "body" && depth === 1;
}

/**
 * The blocks the first container this test names holds, and the text on either side of them.
 * null for a source holding no such container; a caller says in its own words what that means.
 */
export function scanBlocksIn(
  source: string,
  isContainer: ContainerTest
): BlockScan | null {
  const stack: string[] = [];
  const blocks: BlockSlice[] = [];
  let containerDepth: number | null = null;
  let contentStart = -1;
  let sliceStart = -1;

  const atBlockLevel = () =>
    containerDepth !== null && stack.length === containerDepth;
  const takeBlock = (name: string, end: number) => {
    blocks.push({ name, xml: source.slice(sliceStart, end) });
    sliceStart = end;
  };

  let i = 0;
  while (i < source.length) {
    const lt = source.indexOf("<", i);
    if (lt === -1) break;

    const tag = tagAt(source, lt);
    if (tag.kind === "other") {
      i = tag.end;
      continue;
    }

    if (tag.kind === "close") {
      const opened = stack.pop();
      if (opened !== tag.name) {
        throw new DocxImportError(
          "malformed-xml",
          `mismatched tags: ${opened} vs ${tag.name}`
        );
      }
      if (atBlockLevel()) takeBlock(tag.name, tag.end);
      i = tag.end;
      continue;
    }

    if (
      containerDepth === null &&
      tag.kind === "open" &&
      isContainer(tag, stack.length)
    ) {
      stack.push(tag.name);
      containerDepth = stack.length;
      contentStart = tag.end;
      sliceStart = tag.end;
    } else {
      if (atBlockLevel() && tag.kind === "empty") takeBlock(tag.name, tag.end);
      if (tag.kind === "open") stack.push(tag.name);
    }
    i = tag.end;
  }

  if (containerDepth === null) return null;
  if (stack.length > 0) {
    throw new DocxImportError(
      "malformed-xml",
      `elements are left unclosed: ${stack.join(", ")}`
    );
  }
  return {
    prefix: source.slice(0, contentStart),
    blocks,
    suffix: source.slice(sliceStart),
  };
}

export function scanBody(source: string): BlockScan {
  const scan = scanBlocksIn(source, isBodyTag);
  if (scan === null) {
    throw new DocxImportError("missing-body", "document has no w:body");
  }
  return scan;
}
