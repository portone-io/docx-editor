/** Builds the standards-focused long table used only by the browser pagination test. */

import { unzipSync, zipSync } from "fflate";

function tableCell(text: string, merge: "restart" | "continue" | null = null) {
  const mergeXml =
    merge === null
      ? ""
      : merge === "restart"
        ? '<w:vMerge w:val="restart"/>'
        : "<w:vMerge/>";
  return (
    `<w:tc><w:tcPr><w:tcW w:w="3500" w:type="dxa"/>${mergeXml}</w:tcPr>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  );
}

function tableRow(index: number): string {
  const merge = index === 12 ? "restart" : index === 13 ? "continue" : null;
  return (
    '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="780" w:hRule="exact"/></w:trPr>' +
    tableCell(merge === "continue" ? "" : `Row ${index}`, merge) +
    tableCell(`Value ${index}`) +
    "</w:tr>"
  );
}

export function longTableFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const rows = Array.from({ length: 40 }, (_, index) =>
    tableRow(index + 1)
  ).join("");
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    "<w:body><w:tbl>" +
    '<w:tblPr><w:tblW w:w="7000" w:type="dxa"/><w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:insideV w:val="single" w:sz="4" w:color="000000"/>' +
    "</w:tblBorders></w:tblPr>" +
    '<w:tblGrid><w:gridCol w:w="3500"/><w:gridCol w:w="3500"/></w:tblGrid>' +
    '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="exact"/><w:tblHeader/></w:trPr>' +
    tableCell("Heading") +
    tableCell("Value") +
    "</w:tr>" +
    rows +
    "</w:tbl>" +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1247" w:bottom="1304" w:left="1361" w:header="567" w:footer="652" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
