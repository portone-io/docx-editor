/**
 * Builds a document whose table carries a locked row control, the way a template marks the rows a
 * server fills in, beside an open field and a locked one in the paragraph above it.
 *
 * Every cell states its shading as `auto`, which is what a template exported from another word
 * processor writes and what draws as a transparent inline background over the cell.
 */

import { unzipSync, zipSync } from "fflate";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const run = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

const cell = (content: string) =>
  '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:fill="auto"/></w:tcPr>' +
  `<w:p>${content}</w:p></w:tc>`;

const sdtPr = (tag: string, id: number, lock: string | null) =>
  `<w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/>` +
  `${lock === null ? "" : `<w:lock w:val="${lock}"/>`}</w:sdtPr>`;

const field = (tag: string, id: number, lock: string | null, value: string) =>
  `<w:sdt>${sdtPr(tag, id, lock)}<w:sdtContent>${run(value)}</w:sdtContent></w:sdt>`;

const lockedRow = (id: number, first: string, second: string) =>
  `<w:sdt>${sdtPr("PRICE_ROWS", id, "sdtContentLocked")}<w:sdtContent>` +
  `<w:tr>${cell(run(first))}${cell(run(second))}</w:tr>` +
  "</w:sdtContent></w:sdt>";

export function lockedRowsFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}"><w:body>` +
    `<w:p>${run("Payee ")}${field("PAYEE", 10, "sdtContentLocked", "Sample Name")}` +
    `${run(" memo ")}${field("MEMO", 11, null, "Open note")}</w:p>` +
    "<w:tbl>" +
    '<w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>' +
    '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>' +
    '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
    '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>' +
    "</w:tblBorders></w:tblPr>" +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(run("Item"))}${cell(run("Fee"))}</w:tr>` +
    lockedRow(21, "Lessons", "Per session 100") +
    lockedRow(22, "Reviews", "Per item 30") +
    `<w:tr>${cell(run("Extra"))}${cell(run("Agreed"))}</w:tr>` +
    "</w:tbl>" +
    `<w:p>${run("Closing")}</w:p>` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1247" w:bottom="1304" w:left="1361" w:header="567" w:footer="652" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
