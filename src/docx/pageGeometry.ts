/**
 * The paper a document is written on: its size and its margins.
 *
 * A document says so in the `w:sectPr` at the end of its body. The values are read for the screen
 * only: the `w:sectPr` itself rides out with the rest of the preserved tail, so what a document
 * declares is never rewritten from these numbers.
 *
 * Everything here is in twips (1/1440 inch), the unit `w:pgSz` and `w:pgMar` are written in.
 * The one derived value the rest of the editor asks for in two units is the body width, which
 * a table writes down in twips and the sheet draws in pixels, so `bodyWidth` hands out both
 * views of the one number rather than letting each caller convert for itself.
 */

import { toNumber, wAttr } from "../ooxml/units";
import { childByLocalName } from "../ooxml/xml";

/** The paper and the margins a section lays down, in twips */
export interface PageGeometry {
  /** The paper width. §17.6.13 `w:pgSz/@w:w` */
  widthTwips: number;
  /** The paper height. §17.6.13 `w:pgSz/@w:h` */
  heightTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  marginTopTwips: number;
  marginBottomTwips: number;
}

const TWIPS_PER_CM = 1440 / 2.54;

/** A twip is 1/1440 inch, and CSS calls an inch 96 pixels */
const PX_PER_TWIP = 96 / 1440;

function cm(value: number): number {
  return Math.round(value * TWIPS_PER_CM);
}

export function twipsToPx(twips: number): number {
  return twips * PX_PER_TWIP;
}

/**
 * The paper a document that says nothing is drawn on: A4 portrait with the margins this
 * editor has always used.
 *
 * A document with no `w:sectPr`, or one whose `w:pgSz` cannot be read, is drawn on this, which
 * is what every document was drawn on before the geometry was read at all.
 */
export const A4_PORTRAIT: PageGeometry = {
  widthTwips: cm(21),
  heightTwips: cm(29.7),
  marginLeftTwips: cm(2.2),
  marginRightTwips: cm(2.2),
  marginTopTwips: cm(2),
  marginBottomTwips: cm(2),
};

/**
 * A page is at least this wide and this tall, and a margin is never negative.
 *
 * `w:pgSz` is `ST_TwipsMeasure`, which an unsigned type leaves free to hold anything up to
 * 4294967295, and a margin is signed. A document naming a paper a millimetre wide, or margins
 * wider than the paper, would leave no body to draw on at all, so a size that cannot hold a
 * body falls back rather than drawing a sheet nothing fits on.
 */
const MIN_PAGE_TWIPS = cm(1);
const MAX_PAGE_TWIPS = cm(600);

function readSize(value: number | null): number | null {
  if (value === null) return null;
  if (!Number.isFinite(value)) return null;
  if (value < MIN_PAGE_TWIPS || value > MAX_PAGE_TWIPS) return null;
  return Math.round(value);
}

function readMargin(value: number | null, fallback: number): number {
  if (value === null || !Number.isFinite(value)) return fallback;
  // A negative margin puts the body off the paper. Word writes them for gutters we do not draw
  if (value < 0) return 0;
  return Math.min(Math.round(value), MAX_PAGE_TWIPS);
}

/**
 * The geometry a `w:sectPr` element lays down, with anything it leaves unsaid taken from A4.
 *
 * `w:orient` is not applied on top of the size: Word writes the width and the height already
 * swapped for a landscape section, so honouring both would turn the paper back upright.
 */
export function readPageGeometry(sectPr: Element | null): PageGeometry {
  if (!sectPr) return A4_PORTRAIT;

  const pgSz = childByLocalName(sectPr, "pgSz");
  const pgMar = childByLocalName(sectPr, "pgMar");

  const width = pgSz ? readSize(toNumber(wAttr(pgSz, "w"))) : null;
  const height = pgSz ? readSize(toNumber(wAttr(pgSz, "h"))) : null;
  // A size only half readable is not a paper. Both come from the document or neither does
  const readable = width !== null && height !== null;

  const marginLeft = pgMar ? toNumber(wAttr(pgMar, "left")) : null;
  const marginRight = pgMar ? toNumber(wAttr(pgMar, "right")) : null;
  const marginTop = pgMar ? toNumber(wAttr(pgMar, "top")) : null;
  const marginBottom = pgMar ? toNumber(wAttr(pgMar, "bottom")) : null;

  const geometry: PageGeometry = {
    widthTwips: readable ? width : A4_PORTRAIT.widthTwips,
    heightTwips: readable ? height : A4_PORTRAIT.heightTwips,
    marginLeftTwips: readMargin(marginLeft, A4_PORTRAIT.marginLeftTwips),
    marginRightTwips: readMargin(marginRight, A4_PORTRAIT.marginRightTwips),
    marginTopTwips: readMargin(marginTop, A4_PORTRAIT.marginTopTwips),
    marginBottomTwips: readMargin(marginBottom, A4_PORTRAIT.marginBottomTwips),
  };

  // Margins wider than the paper leave no body. Such a section is drawn on A4 instead
  return hasBody(geometry) ? geometry : A4_PORTRAIT;
}

function hasBody(geometry: PageGeometry): boolean {
  return bodyWidthTwips(geometry) > 0 && bodyHeightTwips(geometry) > 0;
}

/** The width one line of body text occupies: the paper less the side margins */
export function bodyWidthTwips(geometry: PageGeometry): number {
  return (
    geometry.widthTwips - geometry.marginLeftTwips - geometry.marginRightTwips
  );
}

/** The height the body occupies on one page: the paper less the top and bottom margins */
export function bodyHeightTwips(geometry: PageGeometry): number {
  return (
    geometry.heightTwips - geometry.marginTopTwips - geometry.marginBottomTwips
  );
}

/**
 * The width of one line of body text, in both units the editor fits things to.
 *
 * A table writes its width down in twips (`w:tblW`, `w:tcW`) and the sheet draws an image in
 * pixels, so the width a new table is divided into, the width a column drag stops at and the
 * width an oversized image is shrunk to are all one number in two views.
 */
export interface BodyWidth {
  /** dxa, the unit a table width is written in */
  twips: number;
  /** CSS pixels, the unit the sheet is drawn in */
  px: number;
}

export function bodyWidth(geometry: PageGeometry): BodyWidth {
  const twips = bodyWidthTwips(geometry);
  return { twips, px: twipsToPx(twips) };
}

/**
 * The body width of the paper a document that names none is drawn on.
 *
 * It is what a caller with no document to ask still gets: 11906 - 1247 * 2 = 9412 dxa.
 */
export const A4_BODY_WIDTH: BodyWidth = bodyWidth(A4_PORTRAIT);

/**
 * The geometry of the first section of an already parsed body.
 *
 * A document may hold a `w:sectPr` per section, each with a paper of its own. The screen draws
 * one paper for the whole document, so the first section decides it; the rest ride out
 * untouched in their preserved blocks and are drawn on the first section's paper.
 * `docs/features.md` records that limit.
 *
 * The reading is deliberately forgiving: this runs against a document already opened, and a
 * `w:sectPr` we cannot read is a reason to draw A4, never a reason to refuse the file.
 */
export function readBodyGeometry(body: Element): PageGeometry {
  return readPageGeometry(firstSectPrIn(body));
}

function firstSectPrIn(root: Element): Element | null {
  const found = root.getElementsByTagName("w:sectPr").item(0);
  if (found) return found;
  for (const el of root.getElementsByTagName("*")) {
    if (el.localName === "sectPr") return el;
  }
  return null;
}
