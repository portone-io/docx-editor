/**
 * The inline picture: the values an image node carries, and the DrawingML those values
 * are read out of and written back into.
 *
 * Only the one shape a docx uses for a plain inline picture is interpreted: a
 * `w:drawing` holding a `wp:inline` whose graphic is a `pic:pic` with an embedded blip.
 * A floating anchor, a chart or a diagram, a blip linked to a file outside the package,
 * an image kind a browser cannot draw - all of those are left to the preservation path
 * that was already there, so nothing we cannot draw is ever rewritten.
 *
 * An imported image holds on to its whole original `<w:drawing>` XML, so an untouched
 * image goes back out byte for byte. A resize rewrites nothing but the two extents.
 *
 * `emuToPx` is imported on its own by a builder working in pixels, so nothing here reads
 * `NAMESPACES` at the top level: a top-level read is an impure statement a bundler keeps,
 * and it would hold the naming layer in a bundle that wanted one multiplication.
 */

import { NAMESPACES } from "./names";
import { parseAttrs, readTag } from "./tagScan";
import { childByLocalName, escapeXml, localPart, R_NS } from "./xml";

/**
 * English Metric Units, the unit every DrawingML length is written in.
 * At 96dpi one CSS pixel is 9525 of them.
 */
export const EMU_PER_PX = 9525;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/**
 * The size a drawing is shown at.
 *
 * It is kept in EMU rather than in pixels because EMU is the document's own unit:
 * reading it and writing it back is exact, so an image nobody resized goes out carrying
 * the very numbers it came in with. Pixels are derived from it for the screen only.
 */
export interface ImageExtent {
  cx: number;
  cy: number;
}

/**
 * The image kinds we both draw on screen and write back into a package.
 *
 * A document may embed a tiff or an emf as well; we do not interpret those, because a
 * browser cannot draw them and a broken picture on screen is worse than the preserved
 * placeholder.
 */
export const IMAGE_MIMES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/bmp",
] as const;

export type ImageMime = (typeof IMAGE_MIMES)[number];

/**
 * A data URL carrying one of the kinds above.
 *
 * The src leaves for the screen as an `img` attribute and comes back in again, so only
 * this one shape is let through: a value tampered with along the way can neither point
 * the browser somewhere else (`javascript:`, an external host) nor break out of the
 * attribute, because base64 has no quotes in it.
 */
const IMAGE_DATA_URL = new RegExp(
  `^data:(?:${IMAGE_MIMES.join("|")});base64,[A-Za-z0-9+/]+={0,2}$`
);

export function isImageMime(value: unknown): value is ImageMime {
  return IMAGE_MIMES.some((mime) => mime === value);
}

/** The image bytes as a data URL. null for anything else */
export function toImageSrc(value: unknown): string | null {
  return typeof value === "string" && IMAGE_DATA_URL.test(value) ? value : null;
}

/** The kind a src declares. It has already been checked by `toImageSrc` */
export function imageMimeOf(src: string): ImageMime | null {
  const mime = src.slice("data:".length, src.indexOf(";"));
  return isImageMime(mime) ? mime : null;
}

/** The base64 payload of a src. It has already been checked by `toImageSrc` */
export function imageBase64Of(src: string): string {
  return src.slice(src.indexOf(",") + 1);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/**
 * A length written in EMU. Only whole positive numbers get through, because these values
 * end up in a size attribute and in a CSS length
 */
function positiveInt(value: unknown): number | null {
  const parsed = typeof value === "string" ? Number.parseInt(value, 10) : value;
  return typeof parsed === "number" && Number.isInteger(parsed) && parsed > 0
    ? parsed
    : null;
}

export function toImageExtent(value: unknown): ImageExtent | null {
  if (!isRecord(value)) return null;
  const cx = positiveInt(value.cx);
  const cy = positiveInt(value.cy);
  return cx === null || cy === null ? null : { cx, cy };
}

/** What one inline picture states about itself */
export interface DrawingPicture {
  /** The relationship id the image bytes hang off */
  relId: string;
  extent: ImageExtent;
  /** The alternative text, from `wp:docPr descr`. null when there is none */
  alt: string | null;
}

function childOf(el: Element | null, name: string): Element | null {
  return el ? childByLocalName(el, name) : null;
}

function relAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(R_NS, name) ?? el.getAttribute(`r:${name}`);
}

function readExtent(el: Element | null): ImageExtent | null {
  if (!el) return null;
  return toImageExtent({
    cx: el.getAttribute("cx"),
    cy: el.getAttribute("cy"),
  });
}

/**
 * Reads a `w:drawing` as an inline picture. null for a drawing we do not interpret,
 * which leaves the caller on the preservation path.
 *
 * Children we do not know about (`wp14` extensions, a crop, a rotation) do not stop us,
 * because the original XML is kept whole and goes back out as it came.
 */
export function readDrawingPicture(drawing: Element): DrawingPicture | null {
  // An anchored drawing carries wrapping and page positioning we do not model
  const inline = childOf(drawing, "inline");
  if (!inline) return null;

  const extent = readExtent(childOf(inline, "extent"));
  if (!extent) return null;

  const graphicData = childOf(childOf(inline, "graphic"), "graphicData");
  // A chart or a diagram sits under a different uri and is not a picture at all
  if (!graphicData || graphicData.getAttribute("uri") !== NAMESPACES.pic) {
    return null;
  }

  const blip = childOf(
    childOf(childOf(graphicData, "pic"), "blipFill"),
    "blip"
  );
  if (!blip) return null;
  // A linked blip keeps its bytes outside the package, so there is nothing here to read
  if (relAttr(blip, "link") !== null) return null;
  const relId = relAttr(blip, "embed");
  if (relId === null) return null;

  const descr = childOf(inline, "docPr")?.getAttribute("descr") ?? null;
  return { relId, extent, alt: descr === "" ? null : descr };
}

/**
 * Follow the same direct-child paths as the picture reader. An extension can contain arbitrary
 * XML (ECMA-376 Part 4, dml-main.xsd CT_OfficeArtExtension), including its own transform or size,
 * so neither the first transform nor every element called extent identifies the picture's size.
 */
function isPictureSize(path: readonly string[]): boolean {
  const names = path.join("/");
  return (
    names === "drawing/inline/extent" ||
    names === "drawing/inline/graphic/graphicData/pic/spPr/xfrm/ext"
  );
}

/** Change only the size values, preserving quoting, namespace declarations and other attributes. */
function resizeTag(open: string, extent: ImageExtent): string {
  const tag = readTag(open, 0);
  if (!tag) return open;
  const attrs = new Map(
    parseAttrs(
      open.slice(tag.nameEnd, tag.end - (tag.kind === "empty" ? 2 : 1))
    ) ?? []
  );
  return open.replace(
    /(\s+)([^\s=/>]+)(\s*=\s*)(["'])([\s\S]*?)\4/g,
    (
      whole: string,
      gap: string,
      name: string,
      equals: string,
      quote: string
    ) => {
      if (name !== "cx" && name !== "cy") return whole;
      const size = extent[name];
      if (Number(attrs.get(name)) === size) return whole;
      return `${gap}${name}${equals}${quote}${size}${quote}`;
    }
  );
}

/**
 * The same drawing XML with its inline extent and its picture transform's extent set to this size.
 *
 * Only those two elements' size attributes are touched; every other byte stays where it was.
 * Values already equal to the requested size keep their original spelling as well.
 */
export function withExtent(xml: string, extent: ImageExtent): string {
  const path: string[] = [];
  const parts: string[] = [];
  let at = 0;
  let kept = 0;
  for (;;) {
    const lt = xml.indexOf("<", at);
    if (lt === -1) break;
    const tag = readTag(xml, lt);
    if (!tag) return xml;
    if (tag.kind === "close") {
      path.pop();
    } else if (tag.kind !== "other") {
      path.push(localPart(tag.name));
      if (isPictureSize(path)) {
        parts.push(
          xml.slice(kept, lt),
          resizeTag(xml.slice(lt, tag.end), extent)
        );
        kept = tag.end;
      }
      if (tag.kind === "empty") path.pop();
    }
    at = tag.end;
  }
  parts.push(xml.slice(kept));
  return parts.join("");
}

/** A picture that was inserted during editing and has no original XML to go back to */
export interface NewImage {
  relId: string;
  /** The drawing's id within the document. Word shows it in the selection pane */
  docPrId: number;
  extent: ImageExtent;
  alt: string | null;
}

/**
 * The smallest `w:drawing` Word opens as an inline picture.
 *
 * Every prefix used is declared right where it is used. The document element declares
 * some of them already, but not all producers declare `a` and `pic`, and a duplicate
 * declaration of the same namespace costs nothing.
 */
export function imageDrawingXml(image: NewImage): string {
  const size = `cx="${image.extent.cx}" cy="${image.extent.cy}"`;
  const name = `Picture ${image.docPrId}`;
  const descr = image.alt === null ? "" : ` descr="${escapeXml(image.alt)}"`;
  return (
    "<w:drawing>" +
    `<wp:inline xmlns:wp="${NAMESPACES.wp}" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent ${size}/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${image.docPrId}" name="${name}"${descr}/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NAMESPACES.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${NAMESPACES.a}">` +
    `<a:graphicData uri="${NAMESPACES.pic}">` +
    `<pic:pic xmlns:pic="${NAMESPACES.pic}">` +
    `<pic:nvPicPr><pic:cNvPr id="${image.docPrId}" name="${name}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${escapeXml(image.relId)}"/>` +
    "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${size}/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
  );
}
