/**
 * Rebuilds an edited paragraph back into OOXML.
 *
 * Body paragraphs and paragraphs inside table cells take the same path.
 *
 * A fragment the reader kept whole goes back where it stood: one kept inside a run is written as
 * a piece of that run, and one kept beside the runs is written between them, which is what the
 * two content models admit.
 *
 * The inlines are grouped from the outside in: neighbours sharing the outermost wrapper each stands
 * inside go back into it, that grouping is made again one wrapper deeper, and at the bottom the
 * neighbours that share their formatting become one run. `schema/wrappers` says which wrapper is
 * which depth, and `docx/wrappers` says what each of them opens and closes as, so a wrapper of a
 * new kind is written here without this file knowing anything about it.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { elementXml, emptyTagXml, openTagXml } from "../ooxml/element";
import { DocxExportError } from "../ooxml/errors";
import {
  imageDrawingXml,
  toImageExtent,
  toImageSrc,
  withExtent,
} from "../ooxml/image";
import { wName } from "../ooxml/names";
import { escapeXml } from "../ooxml/xml";
import { wrapperMarks } from "../schema/wrappers";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import type { ImageRefs } from "./media";
import { wrapperKindOf } from "./wrappers";

/** Whether neighbouring inlines can be grouped together */
function sameMark(a: Mark | null, b: Mark | null): boolean {
  if (a === null || b === null) return a === b;
  return a.eq(b);
}

function markOf(node: PMNode, name: string): Mark | null {
  return node.marks.find((mark) => mark.type.name === name) ?? null;
}

/** The attribute text a node attr holds, which is either what the original wrote or nothing */
export function rawAttrsOf(value: unknown): string | null {
  return typeof value === "string" ? value : null;
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
          ? [
              elementXml(
                wName("t"),
                [["xml:space", "preserve"]],
                [escapeXml(text)]
              ),
            ]
          : [];
        return index < pieces.length - 1
          ? [
              ...rendered,
              emptyTagXml(wName("tab"), rawAttrsOf(tab?.attrs.tabAttrs)),
            ]
          : rendered;
      })
      .join("");
  }
  if (node.type.name === "hardBreak")
    return emptyTagXml(wName("br"), rawAttrsOf(node.attrs.brAttrs));
  // A run child kept whole goes back inside the run it stood in, which is the one place
  // `EG_RunInnerContent` admits it
  if (node.type.name === "rawRunContent") return preservedXml(node);
  if (node.type.name === "image") return renderImage(node, images);
  if (node.type.name === "commentReference") {
    const original: unknown = node.attrs.referenceXml;
    if (typeof original === "string") return original;
    const id: unknown = node.attrs.id;
    if (typeof id === "string") {
      return elementXml(wName("commentReference"), [[wName("id"), id]]);
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
      return elementXml(wName(name), [[wName("id"), id]]);
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
    parts.push({
      kind: "raw",
      xml: elementXml(wName(name), [[wName("id"), id]]),
    });
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

function renderParagraphPart(part: ParagraphPart): string {
  if (part.kind === "raw") return part.xml;
  const open = openTagXml(wName("r"), rawAttrsOf(part.mark?.attrs.rAttrs));
  const rPr: unknown = part.mark?.attrs.rPr;
  return (
    open +
    (typeof rPr === "string" ? rPr : "") +
    part.pieces.join("") +
    "</w:r>"
  );
}

/** The neighbours that stood inside one and the same wrapper at this depth, or outside any of them */
interface WrapperGroup {
  mark: Mark | null;
  children: PMNode[];
}

/**
 * The inlines cut into the stretches that share a wrapper at this depth.
 *
 * The wrappers are read off the marks by position rather than by the depth they claim, so a link
 * made across a stretch that only partly stands inside a control still groups with what it covers.
 */
function groupsAtDepth(
  children: readonly PMNode[],
  depth: number
): WrapperGroup[] {
  const groups: WrapperGroup[] = [];
  for (const child of children) {
    const mark = wrapperMarks(child)[depth] ?? null;
    const last = groups.at(-1);
    if (last && sameMark(last.mark, mark)) last.children.push(child);
    else groups.push({ mark, children: [child] });
  }
  return groups;
}

/** The runs and the fragments kept whole that these inlines go back out as */
function renderParts(children: readonly PMNode[], images: ImageRefs): string {
  const parts: ParagraphPart[] = [];
  for (const child of children) addInline(parts, child, images);
  return parts.map(renderParagraphPart).join("");
}

/**
 * Puts every wrapper from this depth inwards back around the content it held, and writes the runs
 * once no wrapper is left.
 */
function renderWrapped(
  children: readonly PMNode[],
  depth: number,
  refs: ExportRefs
): string {
  return groupsAtDepth(children, depth)
    .map((group) => {
      if (group.mark === null) return renderParts(group.children, refs.images);
      const kind = wrapperKindOf(group.mark);
      const body = renderWrapped(group.children, depth + 1, refs);
      return kind.open(group.mark, refs) + body + kind.close(group.mark);
    })
    .join("");
}

export function serializeParagraph(
  node: PMNode,
  refs: ExportRefs = NO_EXPORT_REFS
): string {
  const open = openTagXml(wName("p"), rawAttrsOf(node.attrs.pAttrs));
  const pPr: unknown = node.attrs.pPr;
  const body = renderWrapped(node.children, 0, refs);
  return open + (typeof pPr === "string" ? pPr : "") + body + "</w:p>";
}
