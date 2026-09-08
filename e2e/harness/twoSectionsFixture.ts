/**
 * Builds the two-section document used only by the browser pagination test.
 *
 * The first section is A4 upright and the second is A4 on its side, with deeper margins above and
 * below than the first: the pages of the second section can only stand where its own paper puts
 * them, and the sheet, whose padding is the first section's, has to be drawn down to the deeper
 * margin the last page ends on.
 *
 * The paper of each section is exported as the twips it is written from, so the test asserting
 * where the pages land measures against the same numbers the document was built with.
 */

import { unzipSync, zipSync } from "fflate";

/** One paper as `w:pgSz` and `w:pgMar` write it, in twips */
export interface FixturePaper {
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  /** Which way round the size was turned, which a producer writes beside the swapped `w:pgSz` */
  orient: "landscape" | null;
}

/** A4 upright: the paper the sheet is drawn on, since it is the first section's */
export const PORTRAIT_PAPER: FixturePaper = {
  width: 11906,
  height: 16838,
  top: 1134,
  right: 1247,
  bottom: 1304,
  left: 1361,
  orient: null,
};

/** The same paper on its side, with the size already swapped the way Word writes it */
export const LANDSCAPE_PAPER: FixturePaper = {
  width: 16838,
  height: 11906,
  top: 1701,
  right: 1247,
  bottom: 2268,
  left: 1361,
  orient: "landscape",
};

function sectPr(paper: FixturePaper): string {
  const orient = paper.orient === null ? "" : ` w:orient="${paper.orient}"`;
  return (
    "<w:sectPr>" +
    `<w:pgSz w:w="${paper.width}" w:h="${paper.height}"${orient}/>` +
    `<w:pgMar w:top="${paper.top}" w:right="${paper.right}"` +
    ` w:bottom="${paper.bottom}" w:left="${paper.left}"` +
    ' w:header="567" w:footer="652" w:gutter="0"/>' +
    "</w:sectPr>"
  );
}

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

/** The width the table is written at, narrower than the body of either section's paper */
const TABLE_DXA = 8000;

/**
 * A table narrower than either section's body, so a column drag has room to widen it past the
 * width the first section's paper would have allowed
 */
const TABLE =
  "<w:tbl><w:tblPr>" +
  `<w:tblW w:w="${TABLE_DXA}" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
  BORDERS +
  "</w:tblPr>" +
  `<w:tblGrid><w:gridCol w:w="${TABLE_DXA / 2}"/>` +
  `<w:gridCol w:w="${TABLE_DXA / 2}"/></w:tblGrid>` +
  "<w:tr>" +
  `<w:tc><w:tcPr><w:tcW w:w="${TABLE_DXA / 2}" w:type="dxa"/></w:tcPr>` +
  "<w:p><w:r><w:t>Wide left</w:t></w:r></w:p></w:tc>" +
  `<w:tc><w:tcPr><w:tcW w:w="${TABLE_DXA / 2}" w:type="dxa"/></w:tcPr>` +
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
    paragraph(
      "Last portrait paragraph",
      `<w:pPr>${sectPr(PORTRAIT_PAPER)}</w:pPr>`
    ) +
    TABLE +
    sideways +
    sectPr(LANDSCAPE_PAPER) +
    "</w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
