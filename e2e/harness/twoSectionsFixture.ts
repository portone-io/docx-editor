/**
 * Builds the two-section document used only by the browser pagination test.
 *
 * The first section is A4 portrait and the second is A4 landscape, with the same margins on
 * both, so the only thing that can move the page boundaries of the second section is its own
 * paper. The margins are the ones every fixture of this project uses, so the paragraphs are
 * drawn the way the other page tests draw them.
 */

import { unzipSync, zipSync } from "fflate";

const MARGINS =
  '<w:pgMar w:top="1134" w:right="1247" w:bottom="1304" w:left="1361"' +
  ' w:header="567" w:footer="652" w:gutter="0"/>';

/** A4 upright: the paper the sheet is drawn on, since it is the first section's */
export const PORTRAIT_SECT_PR = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${MARGINS}</w:sectPr>`;

/** The same paper on its side, with the size already swapped the way Word writes it */
export const LANDSCAPE_SECT_PR =
  '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
  `${MARGINS}</w:sectPr>`;

function paragraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t>${text}</w:t></w:r></w:p>`;
}

const BORDERS =
  "<w:tblBorders>" +
  '<w:top w:val="single" w:sz="4" w:color="000000"/>' +
  '<w:left w:val="single" w:sz="4" w:color="000000"/>' +
  '<w:bottom w:val="single" w:sz="4" w:color="000000"/>' +
  '<w:right w:val="single" w:sz="4" w:color="000000"/>' +
  '<w:insideH w:val="single" w:sz="4" w:color="000000"/>' +
  '<w:insideV w:val="single" w:sz="4" w:color="000000"/>' +
  "</w:tblBorders>";

/**
 * A table narrower than either section's body, so a column drag has room to widen it past the
 * width the first section's paper would have allowed
 */
const TABLE =
  "<w:tbl><w:tblPr>" +
  '<w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  BORDERS +
  "</w:tblPr>" +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  "<w:tr>" +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  "<w:p><w:r><w:t>Wide left</w:t></w:r></w:p></w:tc>" +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  "<w:p><w:r><w:t>Wide right</w:t></w:r></w:p></w:tc>" +
  "</w:tr></w:tbl>";

export function twoSectionsFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const upright = Array.from({ length: 3 }, (_, index) =>
    paragraph(`Portrait paragraph ${index + 1}`)
  ).join("");
  // Enough to run past the end of a second landscape page, so the page before it is a full one
  const sideways = Array.from({ length: 90 }, (_, index) =>
    paragraph(`Landscape paragraph ${index + 1}`)
  ).join("");
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    "<w:body>" +
    upright +
    // The paragraph that ends a section carries that section's own properties (§17.6.17)
    paragraph("Last portrait paragraph", `<w:pPr>${PORTRAIT_SECT_PR}</w:pPr>`) +
    TABLE +
    sideways +
    LANDSCAPE_SECT_PR +
    "</w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
