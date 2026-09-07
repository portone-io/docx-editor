/**
 * Finds the span each individual body block occupies in the raw document.xml text and slices it out.
 *
 * A block that was not edited is exported by writing this fragment back out untouched,
 * which keeps its original bytes intact.
 */

import { DocxImportError } from "../ooxml/errors";
import { readTag, type Tag } from "../ooxml/tagScan";
import { localPart } from "../ooxml/xml";

export interface BodyBlockSlice {
  /** The tag name exactly as written (e.g. "w:p", "w:tbl", "w:sectPr") */
  name: string;
  /** The original XML fragment. It also carries the whitespace that sat between this block and the one before it */
  xml: string;
}

export interface BodyScan {
  /** Everything in the raw text before the first block */
  prefix: string;
  blocks: BodyBlockSlice[];
  /** Everything in the raw text after the last block */
  suffix: string;
}

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

function isBodyTag(tag: Tag, depth: number): boolean {
  return localPart(tag.name) === "body" && depth === 1 && tag.kind === "open";
}

export function scanBody(source: string): BodyScan {
  const stack: string[] = [];
  const blocks: BodyBlockSlice[] = [];
  let bodyDepth: number | null = null;
  let contentStart = -1;
  let sliceStart = -1;

  const atBodyLevel = () => bodyDepth !== null && stack.length === bodyDepth;
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
      if (atBodyLevel()) takeBlock(tag.name, tag.end);
      i = tag.end;
      continue;
    }

    if (bodyDepth === null && isBodyTag(tag, stack.length)) {
      stack.push(tag.name);
      bodyDepth = stack.length;
      contentStart = tag.end;
      sliceStart = tag.end;
    } else {
      if (atBodyLevel() && tag.kind === "empty") takeBlock(tag.name, tag.end);
      if (tag.kind === "open") stack.push(tag.name);
    }
    i = tag.end;
  }

  if (bodyDepth === null) {
    throw new DocxImportError("missing-body", "document has no w:body");
  }
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
