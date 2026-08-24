/**
 * Finds the span each individual body block occupies in the raw document.xml text and slices it out.
 *
 * A block that was not edited is exported by writing this fragment back out untouched,
 * which keeps its original bytes intact.
 */

import { DocxImportError } from "../ooxml/errors";
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

interface OpenTag {
  name: string;
  end: number;
  selfClosing: boolean;
}

interface CloseTag {
  name: string;
  end: number;
}

function skipPast(source: string, from: number, marker: string): number {
  const at = source.indexOf(marker, from);
  if (at === -1) {
    throw new DocxImportError("malformed-xml", `no ${marker} to close a tag`);
  }
  return at + marker.length;
}

/** Skips a whole piece of non-tag syntax (declaration, comment, CDATA). null if it is a tag */
function skipNonTag(source: string, lt: number): number | null {
  if (source.startsWith("<?", lt)) return skipPast(source, lt, "?>");
  if (source.startsWith("<!--", lt)) return skipPast(source, lt, "-->");
  if (source.startsWith("<![CDATA[", lt)) return skipPast(source, lt, "]]>");
  if (source.startsWith("<!", lt)) return skipPast(source, lt, ">");
  return null;
}

/** Reads a single opening tag. A `>` inside quotes is not treated as the end of the tag */
function readOpenTag(source: string, lt: number): OpenTag {
  let i = lt + 1;
  const nameStart = i;
  while (i < source.length && !" \t\r\n/>".includes(source[i])) i += 1;
  const name = source.slice(nameStart, i);
  if (!name)
    throw new DocxImportError("malformed-xml", "a tag has an empty name");

  let quote: string | null = null;
  while (i < source.length) {
    const ch = source[i];
    if (quote !== null) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      return { name, end: i + 1, selfClosing: source[i - 1] === "/" };
    }
    i += 1;
  }
  throw new DocxImportError("malformed-xml", "a tag is left unclosed");
}

function readCloseTag(source: string, lt: number): CloseTag {
  const gt = source.indexOf(">", lt);
  if (gt === -1)
    throw new DocxImportError("malformed-xml", "a tag is left unclosed");
  return { name: source.slice(lt + 2, gt).trim(), end: gt + 1 };
}

function isBodyTag(tag: OpenTag, depth: number): boolean {
  return localPart(tag.name) === "body" && depth === 1 && !tag.selfClosing;
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

    const skipped = skipNonTag(source, lt);
    if (skipped !== null) {
      i = skipped;
      continue;
    }

    if (source.startsWith("</", lt)) {
      const tag = readCloseTag(source, lt);
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

    const tag = readOpenTag(source, lt);
    if (bodyDepth === null && isBodyTag(tag, stack.length)) {
      stack.push(tag.name);
      bodyDepth = stack.length;
      contentStart = tag.end;
      sliceStart = tag.end;
    } else {
      if (atBodyLevel() && tag.selfClosing) takeBlock(tag.name, tag.end);
      if (!tag.selfClosing) stack.push(tag.name);
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
