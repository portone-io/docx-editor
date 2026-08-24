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
 */

import { childByLocalName, escapeXml, R_NS } from "./xml";

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

const PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";

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
  if (!graphicData || graphicData.getAttribute("uri") !== PICTURE_URI) {
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

/** `wp:extent`, the size the drawing takes up in the line */
const EXTENT_TAG = /<([\w.-]+:)?extent\b[^>]*\/>/g;

/** `a:ext` inside `pic:spPr`, the size of the picture frame itself */
const EXT_TAG = /<([\w.-]+:)?ext\b[^>]*\/>/g;

/**
 * The same drawing XML with both extents set to this size.
 *
 * Only those two elements are touched; every other byte of the original stays where it
 * was. Handed the size it was imported with, this gives back the original string
 * unchanged, so nothing has to keep track of whether a resize happened.
 */
export function withExtent(xml: string, extent: ImageExtent): string {
  const size = `cx="${extent.cx}" cy="${extent.cy}"`;
  return xml
    .replace(EXTENT_TAG, (_match, prefix) => `<${prefix ?? ""}extent ${size}/>`)
    .replace(EXT_TAG, (_match, prefix) => `<${prefix ?? ""}ext ${size}/>`);
}

const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
    `<wp:inline xmlns:wp="${WP_NS}" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent ${size}/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${image.docPrId}" name="${name}"${descr}/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A_NS}">` +
    `<a:graphicData uri="${PICTURE_URI}">` +
    `<pic:pic xmlns:pic="${PICTURE_URI}">` +
    `<pic:nvPicPr><pic:cNvPr id="${image.docPrId}" name="${name}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${escapeXml(image.relId)}"/>` +
    "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${size}/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
  );
}
