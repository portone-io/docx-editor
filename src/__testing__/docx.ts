/**
 * Reading fixtures, building tiny docx files, and inspecting edit results, all shared
 * across the tests.
 */

import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { exportDocx } from "../docx/exportDocx";
import type { PageGeometry } from "../docx/pageGeometry";
import type { SessionStore } from "../docx/session";
import type { DocxExportErrorCode, DocxImportErrorCode } from "../ooxml/errors";
import { DocxExportError, DocxImportError } from "../ooxml/errors";

const fixturesDir = join(
  dirname(fileURLToPath(import.meta.url)),
  "../../__fixtures__"
);

export const fixtureNames = readdirSync(fixturesDir).filter((name) =>
  name.endsWith(".docx")
);

export function readFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

/** The one fixture written on paper that is not A4 */
export const LETTER_FIXTURE = "letter-page.docx";

/**
 * US Letter as Word writes it: 8.5in x 11in with an inch of margin all round.
 *
 * The paper `LETTER_FIXTURE` names, kept in one place so that a test working on a geometry and
 * a test reading the paper off a document cannot disagree about what Letter is.
 */
export const LETTER_GEOMETRY: PageGeometry = {
  widthTwips: 12240,
  heightTwips: 15840,
  marginLeftTwips: 1440,
  marginRightTwips: 1440,
  marginTopTwips: 1440,
  marginBottomTwips: 1440,
};

/** The section that names that paper, as the last child of a body */
export const LETTER_SECT_PR =
  "<w:sectPr>" +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr>";

export function decode(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((byte, i) => byte === b[i]);
}

const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function relationships(entries: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    entries +
    "</Relationships>"
  );
}

const PACKAGE_RELS = relationships(
  `<Relationship Id="rId1" Target="word/document.xml" Type="${REL_BASE}/officeDocument"/>`
);

const DOCUMENT_RELS = relationships(
  `<Relationship Id="rId1" Target="styles.xml" Type="${REL_BASE}/styles"/>`
);

const W_NS_DECL =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * Declares the namespaces the fixtures use as well, so a fragment lifted from a real
 * document can be dropped in as is
 */
const BODY_NS_DECL =
  `${W_NS_DECL} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
  ` xmlns:r="${REL_BASE}"`;

/**
 * Takes a body and the contents of styles.xml and zips them into the smallest possible
 * docx
 */
function buildDocx(body: string, styles: string | null): Uint8Array {
  const encoder = new TextEncoder();
  const parts: Record<string, Uint8Array> = {
    "_rels/.rels": encoder.encode(PACKAGE_RELS),
    "word/document.xml": encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<w:document ${BODY_NS_DECL}><w:body>${body}</w:body></w:document>`
    ),
  };
  if (styles !== null) {
    parts["word/_rels/document.xml.rels"] = encoder.encode(DOCUMENT_RELS);
    parts["word/styles.xml"] = encoder.encode(
      `<w:styles ${W_NS_DECL}>${styles}</w:styles>`
    );
  }
  return zipSync(parts);
}

/**
 * The smallest possible docx with only the body swapped in. Given an rPrDefault, a
 * styles.xml is included as well
 */
export function makeDocx(body: string, rPrDefault?: string): Uint8Array {
  return buildDocx(
    body,
    rPrDefault === undefined
      ? null
      : `<w:docDefaults><w:rPrDefault>${rPrDefault}</w:rPrDefault></w:docDefaults>`
  );
}

/** A small docx complete with a styles.xml holding a chain of styles */
export function makeStyledDocx(body: string, styles: string): Uint8Array {
  return buildDocx(body, styles);
}

/** A numbering.xml that defines just one single-level numbered list */
export const ONE_LIST_NUMBERING =
  `<w:numbering ${W_NS_DECL}>` +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
  '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

/** A small docx that includes a numbering.xml, so a new list can be started in it */
export function makeNumberedDocx(
  body: string,
  numberingXml: string = ONE_LIST_NUMBERING
): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(
      `<Relationship Id="rId1" Target="numbering.xml" Type="${REL_BASE}/numbering"/>`
    )
  );
  parts["word/numbering.xml"] = encoder.encode(numberingXml);
  return zipSync(parts);
}

/**
 * A small docx whose body part relates each of these addresses as an external hyperlink
 * relationship, under the id it is keyed by. That is where a `w:hyperlink r:id` finds its address.
 */
export function makeLinkedDocx(
  body: string,
  targets: Readonly<Record<string, string>>
): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(
      Object.entries(targets)
        .map(
          ([id, target]) =>
            `<Relationship Id="${id}" Type="${REL_BASE}/hyperlink"` +
            ` Target="${target}" TargetMode="External"/>`
        )
        .join("")
    )
  );
  return zipSync(parts);
}

export const NOTE_BODY =
  '<w:p><w:r><w:t xml:space="preserve">Text</w:t></w:r>' +
  '<w:r><w:footnoteReference w:id="2"/></w:r>' +
  '<w:r><w:t xml:space="preserve"> and more</w:t></w:r>' +
  '<w:r><w:endnoteReference w:id="3"/></w:r></w:p>';

/** A small package with one regular footnote, one regular endnote, and separator notes. */
export function makeNotesDocx(body: string = NOTE_BODY): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(
      `<Relationship Id="rId4" Target="footnotes.xml" Type="${REL_BASE}/footnotes"/>` +
        `<Relationship Id="rId5" Target="endnotes.xml" Type="${REL_BASE}/endnotes"/>`
    )
  );
  parts["word/footnotes.xml"] = encoder.encode(
    `<w:footnotes ${W_NS_DECL}>` +
      '<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>' +
      '<w:r><w:t xml:space="preserve">Footnote body</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">Second line</w:t></w:r></w:p></w:footnote>' +
      "</w:footnotes>"
  );
  parts["word/endnotes.xml"] = encoder.encode(
    `<w:endnotes ${W_NS_DECL}>` +
      '<w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
      '<w:endnote w:id="3"><w:p><w:r><w:endnoteRef/></w:r>' +
      '<w:r><w:t xml:space="preserve">Endnote body</w:t></w:r></w:p></w:endnote>' +
      "</w:endnotes>"
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
      '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

/** A package with all first-section header/footer variants and page-number fields. */
export function makeHeadersFootersDocx(): Uint8Array {
  const encoder = new TextEncoder();
  const body =
    "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
    '<w:sectPr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:headerReference w:type="default" r:id="rId10"/>' +
    '<w:headerReference w:type="first" r:id="rId11"/>' +
    '<w:headerReference w:type="even" r:id="rId12"/>' +
    '<w:footerReference w:type="default" r:id="rId13"/>' +
    '<w:footerReference w:type="first" r:id="rId14"/>' +
    '<w:footerReference w:type="even" r:id="rId15"/>' +
    '<w:pgNumType w:start="4"/><w:titlePg/>' +
    "</w:sectPr>";
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(
      `<Relationship Id="rId10" Target="header1.xml" Type="${REL_BASE}/header"/>` +
        `<Relationship Id="rId11" Target="header2.xml" Type="${REL_BASE}/header"/>` +
        `<Relationship Id="rId12" Target="header3.xml" Type="${REL_BASE}/header"/>` +
        `<Relationship Id="rId13" Target="footer1.xml" Type="${REL_BASE}/footer"/>` +
        `<Relationship Id="rId14" Target="footer2.xml" Type="${REL_BASE}/footer"/>` +
        `<Relationship Id="rId15" Target="footer3.xml" Type="${REL_BASE}/footer"/>` +
        `<Relationship Id="rId16" Target="settings.xml" Type="${REL_BASE}/settings"/>`
    )
  );
  parts["word/header1.xml"] = encoder.encode(
    `<w:hdr ${W_NS_DECL}><w:p><w:r><w:t xml:space="preserve">Default </w:t></w:r>` +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
      '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> NUMPAGES \\* MERGEFORMAT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:p></w:hdr>"
  );
  parts["word/header2.xml"] = encoder.encode(
    `<w:hdr ${W_NS_DECL}><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>`
  );
  parts["word/header3.xml"] = encoder.encode(
    `<w:hdr ${W_NS_DECL}><w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>`
  );
  parts["word/footer1.xml"] = encoder.encode(
    `<w:ftr ${W_NS_DECL}><w:p><w:r><w:t>Default footer</w:t></w:r></w:p></w:ftr>`
  );
  parts["word/footer2.xml"] = encoder.encode(
    `<w:ftr ${W_NS_DECL}><w:p><w:r><w:t>First footer</w:t></w:r></w:p></w:ftr>`
  );
  parts["word/footer3.xml"] = encoder.encode(
    `<w:ftr ${W_NS_DECL}><w:p><w:r><w:t>Even footer</w:t></w:r></w:p></w:ftr>`
  );
  parts["word/settings.xml"] = encoder.encode(
    `<w:settings ${W_NS_DECL}><w:evenAndOddHeaders/></w:settings>`
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/word/footer3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

/**
 * A 1x1 fully transparent PNG, the smallest image a document can carry.
 * Written as base64 because that is the form both a data URL and this literal need
 */
export const TINY_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNg" +
  "AAIAAAUAAen63NgAAAAASUVORK5CYII=";

export const TINY_PNG_DATA_URL = `data:image/png;base64,${TINY_PNG_BASE64}`;

/** The bytes are backed by a plain ArrayBuffer, which is what the zip helpers ask for */
export function decodeBase64(base64: string): Uint8Array<ArrayBuffer> {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let at = 0; at < binary.length; at += 1) {
    bytes[at] = binary.charCodeAt(at);
  }
  return bytes;
}

export const TINY_PNG = decodeBase64(TINY_PNG_BASE64);

/** A second image (a single red pixel), so a test can tell two media parts apart */
export const OTHER_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP4" +
  "z8DwHwAFAAH/VscvDQAAAABJRU5ErkJggg==";

/** The relationship id the image in `makeImageDocx` hangs off */
export const IMAGE_REL_ID = "rId7";

const PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/**
 * Every prefix a drawing uses, declared on the drawing itself.
 * Word declares wp and r on the document element instead, but a fragment that declares
 * its own is just as valid and keeps the test bodies small
 */
const DRAWING_NS_DECL =
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ` xmlns:pic="${PICTURE_URI}"` +
  ` xmlns:r="${REL_BASE}"`;

export interface PictureOptions {
  relId?: string;
  /** The size in EMU. 1905000 x 952500 is 200 x 100 px */
  cx?: number;
  cy?: number;
  descr?: string;
  /** Points at a file outside the package (`r:link`) instead of at a part inside it */
  linked?: boolean;
  /** The graphic uri, so a test can put something that is not a picture inside a drawing */
  uri?: string;
}

function graphicXml(options: PictureOptions): string {
  const { relId = IMAGE_REL_ID, cx = 1905000, cy = 952500 } = options;
  const blipRef = options.linked ? `r:link="${relId}"` : `r:embed="${relId}"`;
  return (
    `<a:graphic><a:graphicData uri="${options.uri ?? PICTURE_URI}">` +
    "<pic:pic>" +
    '<pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
    `<pic:blipFill><a:blip ${blipRef}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic>"
  );
}

/** A `w:drawing` holding an inline picture, the shape the editor draws as an image */
export function inlineDrawingXml(options: PictureOptions = {}): string {
  const { cx = 1905000, cy = 952500, descr } = options;
  return (
    `<w:drawing${DRAWING_NS_DECL}><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="4" name="Picture 1"${descr === undefined ? "" : ` descr="${descr}"`}/>` +
    "<wp:cNvGraphicFramePr/>" +
    graphicXml(options) +
    "</wp:inline></w:drawing>"
  );
}

/** A `w:drawing` holding a floating picture, which the editor leaves on the preservation path */
export function anchoredDrawingXml(options: PictureOptions = {}): string {
  const { cx = 1905000, cy = 952500 } = options;
  return (
    `<w:drawing${DRAWING_NS_DECL}><wp:anchor distT="0" distB="0" behindDoc="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:wrapNone/><wp:docPr id="5" name="Picture 2"/>' +
    graphicXml(options) +
    "</wp:anchor></w:drawing>"
  );
}

/** A run carrying nothing but a drawing, the way Word writes a picture */
export function drawingRun(drawing: string): string {
  return `<w:r><w:rPr><w:noProof/></w:rPr>${drawing}</w:r>`;
}

const CONTENT_TYPE_DEFAULTS =
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>';

function contentTypes(declarePng: boolean): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    CONTENT_TYPE_DEFAULTS +
    (declarePng ? '<Default Extension="png" ContentType="image/png"/>' : "") +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>"
  );
}

/**
 * A small docx that carries one png at `word/media/image1.png`, related to the body as
 * `IMAGE_REL_ID`, together with a [Content_Types].xml.
 *
 * `declarePng` decides whether the content types already declare the png extension, so a
 * test can watch the export add the declaration or leave it alone. `bytes` puts the media
 * part there as it is, which is what a test needing a large one asks for.
 */
export function makeImageDocx(
  body: string,
  options: {
    base64?: string;
    bytes?: Uint8Array<ArrayBuffer>;
    declarePng?: boolean;
  } = {}
): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeDocx(body));
  parts["word/media/image1.png"] =
    options.bytes ?? decodeBase64(options.base64 ?? TINY_PNG_BASE64);
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(
      `<Relationship Id="${IMAGE_REL_ID}" Target="media/image1.png" Type="${REL_BASE}/image"/>`
    )
  );
  parts["[Content_Types].xml"] = encoder.encode(
    contentTypes(options.declarePng ?? false)
  );
  return zipSync(parts);
}

/**
 * The code the import refusal carries. The code is the stable part of a failure,
 * so tests pin it rather than the wording of the message
 */
export function importErrorCode(run: () => unknown): DocxImportErrorCode {
  try {
    run();
  } catch (error) {
    if (error instanceof DocxImportError) return error.code;
    throw error;
  }
  throw new Error("the import was expected to be refused");
}

/** The code the export refusal carries */
export function exportErrorCode(run: () => unknown): DocxExportErrorCode {
  try {
    run();
  } catch (error) {
    if (error instanceof DocxExportError) return error.code;
    throw error;
  }
  throw new Error("the export was expected to be refused");
}

/** The main part as rewritten from the edit result */
export function documentXmlOf(doc: PMNode, session: SessionStore): string {
  return decode(unzipSync(exportDocx(doc, session))[session.mainPartPath]);
}

/** The original bytes that have to remain on either side of the edited block */
export function surroundings(
  session: SessionStore,
  index: number
): { head: string; tail: string } {
  const xmlOf = (block: { xml: string }) => block.xml;
  return {
    head:
      session.documentPrefix +
      session.blocks.slice(0, index).map(xmlOf).join(""),
    tail:
      session.blocks
        .slice(index + 1)
        .map(xmlOf)
        .join("") + session.documentSuffix,
  };
}
