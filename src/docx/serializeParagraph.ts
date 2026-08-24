/**
 * Rebuilds an edited paragraph back into OOXML.
 *
 * Body paragraphs and paragraphs inside table cells take the same path.
 *
 * The inlines are grouped three times over: neighbours that share their formatting become one run,
 * the runs that share a hyperlink (`w:hyperlink`) go back inside it, and the links and runs that
 * share a content control (`w:sdt`) go back inside the wrapper that mark carries. The control is
 * the outer of the two wrappers, which is the nesting the mark order records (`schema`).
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import {
  imageDrawingXml,
  toImageExtent,
  toImageSrc,
  withExtent,
} from "../ooxml/image";
import { escapeXml } from "../ooxml/xml";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { type LinkRefs, relIdIn, withRelId } from "./hyperlink";
import type { ImageRefs } from "./media";

/** Whether neighbouring inlines can be grouped together */
function sameMark(a: Mark | null, b: Mark | null): boolean {
  if (a === null || b === null) return a === b;
  return a.eq(b);
}

function markOf(node: PMNode, name: string): Mark | null {
  return node.marks.find((mark) => mark.type.name === name) ?? null;
}

export function openTag(name: string, attrs: unknown): string {
  return typeof attrs === "string" ? `<${name} ${attrs}>` : `<${name}>`;
}

function emptyTag(name: string, attrs: unknown): string {
  return typeof attrs === "string" ? `<${name} ${attrs}/>` : `<${name}/>`;
}

/**
 * The drawing an image node goes out as.
 *
 * An imported image goes back out as its own original XML with the extents set to the
 * size the node now holds. Handed a size nobody changed, that is the original string
 * itself, so an untouched image is byte identical.
 * An image inserted during editing has no original, and gets the smallest drawing Word
 * reads, pointing at the media part the export added for it.
 */
function renderImage(node: PMNode, images: ImageRefs): string {
  const extent = toImageExtent(node.attrs.extent);
  const xml: unknown = node.attrs.xml;
  if (typeof xml === "string") return extent ? withExtent(xml, extent) : xml;

  const src = toImageSrc(node.attrs.src);
  if (!src || !extent) {
    throw new DocxExportError(
      "lost-original",
      "an image carries neither its original XML nor a size and bytes to rebuild it from"
    );
  }
  const relId = images.relIdOf(src);
  if (relId === undefined) {
    throw new DocxExportError(
      "unsupported-content",
      "an inserted image has no media relationship; export it through exportDocx"
    );
  }
  return imageDrawingXml({
    relId,
    docPrId: images.takeDocPrId(),
    extent,
    alt: typeof node.attrs.alt === "string" ? node.attrs.alt : null,
  });
}

function renderInline(node: PMNode, images: ImageRefs): string {
  if (node.isText) {
    const tab = markOf(node, "tab");
    return (node.text ?? "")
      .split("\t")
      .flatMap((text, index, pieces) => {
        const rendered = text
          ? [`<w:t xml:space="preserve">${escapeXml(text)}</w:t>`]
          : [];
        return index < pieces.length - 1
          ? [...rendered, emptyTag("w:tab", tab?.attrs.tabAttrs)]
          : rendered;
      })
      .join("");
  }
  if (node.type.name === "hardBreak")
    return emptyTag("w:br", node.attrs.brAttrs);
  if (node.type.name === "image") return renderImage(node, images);
  if (node.type.name === "commentReference") {
    const original: unknown = node.attrs.referenceXml;
    if (typeof original === "string") return original;
    const id: unknown = node.attrs.id;
    if (typeof id === "string") {
      return `<w:commentReference w:id="${escapeXml(id)}"/>`;
    }
    throw new DocxExportError(
      "lost-original",
      "a comment reference has no id to write"
    );
  }
  if (node.type.name === "noteReference") {
    const original: unknown = node.attrs.referenceXml;
    if (typeof original === "string") return original;
    const id: unknown = node.attrs.id;
    const name =
      node.attrs.kind === "endnote" ? "endnoteReference" : "footnoteReference";
    if (typeof id === "string") {
      return `<w:${name} w:id="${escapeXml(id)}"/>`;
    }
    throw new DocxExportError(
      "lost-original",
      "a note reference has no id to write"
    );
  }
  throw new DocxExportError(
    "unsupported-content",
    `inline node we cannot serialize: ${node.type.name}`
  );
}

type ParagraphPart =
  | { kind: "run"; mark: Mark | null; pieces: string[] }
  | { kind: "raw"; xml: string };

/** The parts that stood inside one and the same hyperlink, or outside any of them */
interface LinkGroup {
  link: Mark | null;
  parts: ParagraphPart[];
}

/** The groups that stood inside one and the same content control, or outside any of them */
interface ParagraphGroup {
  sdt: Mark | null;
  links: LinkGroup[];
}

export function preservedXml(node: PMNode): string {
  const xml: unknown = node.attrs.xml;
  if (typeof xml !== "string") {
    throw new DocxExportError(
      "lost-original",
      "a preserved element has lost its original XML"
    );
  }
  return xml;
}

/** Adds one inline to the parts, joining the run before it when they share the same formatting */
function addInline(
  parts: ParagraphPart[],
  child: PMNode,
  images: ImageRefs
): void {
  if (child.type.name === "rawInline") {
    parts.push({ kind: "raw", xml: preservedXml(child) });
    return;
  }
  if (child.type.name === "commentStart" || child.type.name === "commentEnd") {
    const original: unknown = child.attrs.xml;
    if (typeof original === "string") {
      parts.push({ kind: "raw", xml: original });
      return;
    }
    const id: unknown = child.attrs.id;
    if (typeof id !== "string") {
      throw new DocxExportError(
        "lost-original",
        "a comment range marker has no id to write"
      );
    }
    const name =
      child.type.name === "commentStart"
        ? "commentRangeStart"
        : "commentRangeEnd";
    parts.push({ kind: "raw", xml: `<w:${name} w:id="${escapeXml(id)}"/>` });
    return;
  }
  const piece = renderInline(child, images);
  const mark = markOf(child, "run");
  const last = parts.at(-1);
  if (last?.kind === "run" && sameMark(last.mark, mark)) {
    last.pieces.push(piece);
  } else {
    parts.push({ kind: "run", mark, pieces: [piece] });
  }
}

function splitParagraphGroups(
  paragraph: PMNode,
  refs: ExportRefs
): ParagraphGroup[] {
  const groups: ParagraphGroup[] = [];
  paragraph.forEach((child) => {
    const sdt = markOf(child, "sdt");
    const link = markOf(child, "link");
    let group = groups.at(-1);
    if (!group || !sameMark(group.sdt, sdt)) {
      group = { sdt, links: [] };
      groups.push(group);
    }
    let inLink = group.links.at(-1);
    if (!inLink || !sameMark(inLink.link, link)) {
      inLink = { link, parts: [] };
      group.links.push(inLink);
    }
    addInline(inLink.parts, child, refs.images);
  });
  return groups;
}

function renderParagraphPart(part: ParagraphPart): string {
  if (part.kind === "raw") return part.xml;
  const open = openTag("w:r", part.mark?.attrs.rAttrs);
  const rPr: unknown = part.mark?.attrs.rPr;
  return (
    open +
    (typeof rPr === "string" ? rPr : "") +
    part.pieces.join("") +
    "</w:r>"
  );
}

/**
 * The opening tag of the link these parts stood inside.
 *
 * A link that came in with an address goes out pointing at the relationship that address now lives
 * on, which for a link nobody retargeted is the one it arrived on: the tag is then the very string
 * it came as. A link the editor made has no tag of its own and gets the smallest one Word reads.
 * With no relationships to hand out - a serializer running outside an export - an imported link
 * still goes out as it came, and one made here has nothing to point at.
 */
function openLinkTag(mark: Mark, links: LinkRefs): string {
  const prefix: unknown = mark.attrs.linkPrefix;
  const original = typeof prefix === "string" ? prefix : null;
  const href: unknown = mark.attrs.href;
  // A link naming a bookmark alone carries no address, and its wrapper says where it goes
  if (typeof href !== "string") {
    if (original !== null) return original;
    throw new DocxExportError(
      "lost-original",
      "a hyperlink carries neither an address nor the opening XML it goes back out as"
    );
  }
  const relId = links.relIdOf(
    href,
    original === null ? null : relIdIn(original)
  );
  if (relId === undefined) {
    if (original !== null) return original;
    throw new DocxExportError(
      "unsupported-content",
      "an inserted hyperlink has no relationship to point at; export it through exportDocx"
    );
  }
  return withRelId(original ?? "<w:hyperlink>", relId);
}

/** Puts the hyperlink these parts stood inside back around them */
function renderLinkGroup(group: LinkGroup, links: LinkRefs): string {
  const body = group.parts.map(renderParagraphPart).join("");
  if (!group.link) return body;
  return `${openLinkTag(group.link, links)}${body}</w:hyperlink>`;
}

/** Puts the content control these groups stood inside back around them */
function renderParagraphGroup(group: ParagraphGroup, refs: ExportRefs): string {
  const body = group.links
    .map((inLink) => renderLinkGroup(inLink, refs.links))
    .join("");
  if (!group.sdt) return body;
  const prefix: unknown = group.sdt.attrs.sdtPrefix;
  if (typeof prefix !== "string") {
    throw new DocxExportError(
      "lost-original",
      "a content control has lost the opening XML it goes back out as"
    );
  }
  return `${prefix}<w:sdtContent>${body}</w:sdtContent></w:sdt>`;
}

export function serializeParagraph(
  node: PMNode,
  refs: ExportRefs = NO_EXPORT_REFS
): string {
  const open = openTag("w:p", node.attrs.pAttrs);
  const pPr: unknown = node.attrs.pPr;
  const body = splitParagraphGroups(node, refs)
    .map((group) => renderParagraphGroup(group, refs))
    .join("");
  return open + (typeof pPr === "string" ? pPr : "") + body + "</w:p>";
}
