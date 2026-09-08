/**
 * Reads a side story - a comment's body, a footnote's - the way the body of the document is read.
 *
 * A story is a document of the same schema, built by the same block readers, sliced verbatim by
 * the same scanner. That is what lets a comment keep its bold run, its paragraph style and its
 * second paragraph through an edit: the editor no longer holds a body as the plain text it reads
 * as, but as the blocks it was written in, and writes back the ones nobody touched exactly as they
 * arrived (`./serializeStory`).
 *
 * The stories the package arrived holding live on the session; what the document currently says
 * lives on the document node, under `doc.attrs.stories`, so an edit to one rides a transaction and
 * lands in the history beside every other edit.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import { xmlnsAttr } from "../ooxml/element";
import { DocxImportError } from "../ooxml/errors";
import { NAMESPACES, qualify } from "../ooxml/names";
import { attributeByLocalName, elementChildren, parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { isPreservedNode } from "../schema/preservedFragments";
import {
  STORIES_ATTR,
  type StoryJson,
  type StoryKey,
  type StoryKind,
  storiesOf,
  storyKey,
  storyNodeOf,
} from "../schema/stories";
import { type FormattingContext, styledParagraph } from "./formatting";
import { buildParagraph, type ImportSources } from "./importParagraph";
import { buildPreservedBlock } from "./importPreserved";
import { buildTable } from "./importTable";
import { type BlockScan, scanBlocksIn } from "./scan";
import { blockKey, type ImportedBlock, type SessionIdentity } from "./session";

export type { StoryJson, StoryKey, StoryKind } from "../schema/stories";
export {
  asStoryKey,
  sameStory,
  storiesOf,
  storyKey,
  withoutCommentStories,
} from "../schema/stories";

/** The story this key names as a document node, and null where the document holds none */
export const storyOf = storyNodeOf;

/** One story as the package arrived holding it: where it stands, and the blocks it was written in */
export interface ImportedStory {
  key: StoryKey;
  kind: StoryKind;
  id: string;
  partPath: string;
  /** The whole element as written, which is what a story nobody edited goes back out as */
  xml: string;
  /**
   * The container's opening tag as written, e.g. `<w:comment w:id="4" w:author="A">`. A container
   * that arrived self-closing is written out as a pair, since an edited story needs one
   */
  open: string;
  /** Everything after the last block, the closing tag included */
  close: string;
  /** The blocks between them, sliced verbatim; a block's `srcId` indexes into this list */
  blocks: readonly ImportedBlock[];
  doc: PMNode;
}

/** Where a run of stories of one kind is written down, and what tells one of them from the next */
export interface StoryPart {
  kind: StoryKind;
  partPath: string;
  xml: string;
  /** The local name of the element each story stands in, e.g. `comment` or `footnote` */
  entryName: string;
  /** The attribute naming which story an element is */
  idAttr: string;
}

/** What reading a story takes beyond the text of it */
export interface StoryDeps {
  session: SessionIdentity;
  /**
   * The images and links of the part the story stands in. A story holds neither a comment nor a
   * note of its own - WordprocessingML puts neither there - so those two are read empty.
   */
  sources: ImportSources;
  formatting: FormattingContext;
}

/**
 * Moves a single block into a node.
 *
 * A paragraph always opens editable (`./importParagraph`). What is left over is a table whose rows
 * this reader could not take apart, a range marker standing between blocks, and a block this
 * reader has no reader for at all; each stands as one placeholder naming the original fragment.
 */
export function buildBlock(
  el: Element,
  srcId: string,
  sources: ImportSources,
  context: FormattingContext
): PMNode {
  if (el.localName === "p") return buildParagraph(el, srcId, sources);
  if (el.localName === "tbl") {
    const table = buildTable(el, srcId, sources, context);
    if (table) return table;
  }
  return buildPreservedBlock(el, srcId, "body");
}

/**
 * Folds the style chain into the display values.
 *
 * These values are used for display only, so the original XML fragments are left untouched.
 * A table is left alone: which part of it a cell belongs to is what its table style dresses the
 * paragraphs inside by, and `buildTable` is where that is known.
 */
export function withStyleFormats(
  node: PMNode,
  context: FormattingContext
): PMNode {
  if (node.type === docxSchema.nodes.paragraph) {
    return styledParagraph(node, context);
  }
  if (node.childCount === 0 || node.type === docxSchema.nodes.table) {
    return node;
  }
  const children = node.children.map((child) =>
    withStyleFormats(child, context)
  );
  // If no child changed, do not rebuild the node
  if (children.every((child, i) => child === node.child(i))) return node;
  return node.copy(Fragment.fromArray(children));
}

function malformed(what: string): never {
  throw new DocxImportError("malformed-xml", what);
}

/** An element written as one self-closing tag, which holds nothing and has no closing tag to find */
const SELF_CLOSING = /\/>\s*$/;

/**
 * The blocks the first element of this text holds, and the text on either side of them.
 *
 * A container written as one self-closing tag holds no block, and the two ends it would take one
 * between are written out for it. Nothing is written from them unless the story is edited, since
 * an untouched one goes back out as the bytes it arrived as.
 */
function scanElement(xml: string, el: Element): BlockScan {
  const scan = scanBlocksIn(xml, (_tag, depth) => depth === 0);
  if (scan !== null) return scan;
  if (!SELF_CLOSING.test(xml) || el.children.length > 0) {
    malformed(`no ${el.nodeName} element to read from`);
  }
  return {
    prefix: xml.replace(SELF_CLOSING, ">"),
    blocks: [],
    suffix: `</${el.nodeName}>`,
  };
}

/**
 * The blocks a container holds, matched between the text and the DOM.
 *
 * The two have to agree for the same reason the body's do: a verbatim slice is written back out
 * untouched, so a slice that names a different element than the reader read is a slice that would
 * put someone else's markup where this block stood.
 */
function scanContainer(xml: string, el: Element) {
  const scan = scanElement(xml, el);
  const children = elementChildren(el);
  if (children.length !== scan.blocks.length) {
    malformed(
      `the block count of ${el.nodeName} differs between the scan and the DOM`
    );
  }
  children.forEach((child, at) => {
    if (child.nodeName !== scan.blocks[at].name) {
      malformed(
        `story block names disagree: ${child.nodeName} vs ${scan.blocks[at].name}`
      );
    }
  });
  return { scan, children };
}

/**
 * One story as it stands in the part, with its blocks read and sliced.
 *
 * A container holding no block at all still opens editable, the way a body holding nothing but a
 * section does: the placeholder paragraph names an empty original fragment, so a story nobody
 * edited goes back out as the bytes it arrived as.
 */
export function readStory(
  place: { kind: StoryKind; id: string; partPath: string },
  container: { el: Element; xml: string },
  deps: StoryDeps
): ImportedStory {
  const key = storyKey(place.kind, place.id);
  const { scan, children } = scanContainer(container.xml, container.el);
  const nodes = children.map((child, at) =>
    withStyleFormats(
      buildBlock(
        child,
        blockKey(deps.session, key, at),
        deps.sources,
        deps.formatting
      ),
      deps.formatting
    )
  );
  const blocks: ImportedBlock[] = nodes.map((node, at) => ({
    xml: scan.blocks[at].xml,
    node,
  }));
  if (nodes.length === 0) {
    const placeholder = withStyleFormats(
      docxSchema.nodes.paragraph.create({
        srcId: blockKey(deps.session, key, 0),
      }),
      deps.formatting
    );
    nodes.push(placeholder);
    blocks.push({ xml: "", node: placeholder });
  }
  return {
    key,
    kind: place.kind,
    id: place.id,
    partPath: place.partPath,
    xml: container.xml,
    open: scan.prefix,
    close: scan.suffix,
    blocks,
    doc: docxSchema.nodes.doc.create(null, nodes),
  };
}

/** Every story one part holds, in the order the part writes them */
export function readStories(
  part: StoryPart,
  deps: StoryDeps
): readonly ImportedStory[] {
  const root = parseXml(part.xml).documentElement;
  const scan = scanElement(part.xml, root);
  const children = elementChildren(root);
  if (children.length !== scan.blocks.length) {
    malformed(
      `the entry count of ${part.partPath} differs between the scan and the DOM`
    );
  }
  const stories: ImportedStory[] = [];
  const seen = new Set<string>();
  children.forEach((el, at) => {
    if (el.localName !== part.entryName) return;
    const id = attributeByLocalName(el, part.idAttr);
    if (id === null || seen.has(id)) return;
    seen.add(id);
    // The whitespace a part was laid out with sits ahead of the slice; the story itself begins at
    // its own opening tag, so what the writer puts back is the entry and not the layout around it
    const xml = scan.blocks[at].xml;
    stories.push(
      readStory(
        { kind: part.kind, id, partPath: part.partPath },
        { el, xml: xml.slice(xml.indexOf("<")) },
        deps
      )
    );
  });
  return stories;
}

/** The stories of a document keyed the way the session holds them */
export function storiesByKey(
  stories: Iterable<ImportedStory>
): ReadonlyMap<StoryKey, ImportedStory> {
  return new Map(Array.from(stories, (story) => [story.key, story] as const));
}

/**
 * What writing a story takes of a transaction: the document it stands on, and the one step that
 * changes a document attr.
 *
 * A `Transaction` is what a caller passes, and it is named by shape rather than imported because
 * `src/core.ts` must reach neither the editor nor `prosemirror-state` (`src/core.test.ts`), and
 * this module is on the way there.
 */
interface StoryWriter {
  readonly doc: PMNode;
  setDocAttribute(attr: string, value: unknown): this;
}

/** The document with this story written into it, which is the one way a story changes */
export function setStory<T extends StoryWriter>(
  tr: T,
  key: StoryKey,
  story: PMNode
): T {
  const json: StoryJson = story.toJSON();
  return tr.setDocAttribute(STORIES_ATTR, {
    ...storiesOf(tr.doc),
    [key]: json,
  });
}

/** The document with these stories taken out of it, which is what deleting what they say leaves */
export function withoutStories<T extends StoryWriter>(
  tr: T,
  keys: Iterable<StoryKey>
): T {
  const dropped = new Set<string>(keys);
  const stories = storiesOf(tr.doc);
  const kept = Object.fromEntries(
    Object.entries(stories).filter(([held]) => !dropped.has(held))
  );
  if (Object.keys(kept).length === Object.keys(stories).length) return tr;
  return tr.setDocAttribute(STORIES_ATTR, kept);
}

/**
 * A story holding this text, one paragraph broken by a line break wherever the text is.
 *
 * That is the shape a plain-text body has been written in since before there were stories, so a
 * body written through `updateComment` goes out as the bytes it always did.
 */
export function storyFromText(text: string): PMNode {
  const pieces: PMNode[] = [];
  text.split("\n").forEach((line, at) => {
    if (at > 0) pieces.push(docxSchema.nodes.hardBreak.create());
    if (line.length > 0) pieces.push(docxSchema.text(line));
  });
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create(null, pieces),
  ]);
}

/**
 * What one leaf of a story puts on screen.
 *
 * A break ends a line, and a fragment the editor keeps rather than models says what it stands for
 * - the character a `w:noBreakHyphen` draws, the words inside a field - which is the same answer
 * `docx/importPolicy` gave when a body was flattened on the way in.
 */
function leafText(leaf: PMNode): string {
  if (leaf.type === docxSchema.nodes.hardBreak) return "\n";
  if (!isPreservedNode(leaf)) return "";
  if (leaf.attrs.display === "break") return "\n";
  return typeof leaf.attrs.text === "string" ? leaf.attrs.text : "";
}

/** What a story reads as on screen: its paragraphs joined by newlines, its breaks and tabs kept */
export function storyText(story: PMNode | null): string {
  if (story === null) return "";
  return story.textBetween(0, story.content.size, "\n", leafText);
}

/**
 * The story with the thread key on the last paragraph of it, which is where Word keeps it.
 *
 * The key is written into the paragraph's own attribute text rather than into a model attr,
 * because that text is what the writer puts back and a `w14:paraId` is not a name this schema
 * holds. A paragraph already carrying one keeps it: a key already written is what the thread
 * state elsewhere is keyed by, so re-pointing it would orphan that state.
 */
export function withThreadKeyOn(story: PMNode, paraId: string): PMNode {
  let at = -1;
  story.forEach((block, _offset, index) => {
    if (block.type === docxSchema.nodes.paragraph) at = index;
  });
  if (at === -1) return story;
  const paragraph = story.child(at);
  const written: unknown = paragraph.attrs.pAttrs;
  const pAttrs = typeof written === "string" ? written : "";
  if (/\bw14:paraId\s*=/.test(pAttrs)) return story;
  const declared = /\bxmlns:w14\s*=/.test(pAttrs)
    ? ""
    : `${xmlnsAttr("w14")[0]}="${NAMESPACES.w14}" `;
  const keyed = paragraph.type.create(
    {
      ...paragraph.attrs,
      pAttrs: `${pAttrs === "" ? "" : `${pAttrs} `}${declared}${qualify("w14", "paraId")}="${paraId}"`,
    },
    paragraph.content,
    paragraph.marks
  );
  return story.copy(story.content.replaceChild(at, keyed));
}
