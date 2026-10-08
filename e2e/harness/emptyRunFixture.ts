/**
 * Builds a form whose blanks are runs with a yellow highlight and no text, the way a server writing
 * the form leaves them: in the empty cells of a table, after a label, and as all an inline content
 * control holds, open or locked.
 */

import { unzipSync, zipSync } from "fflate";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const run = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

/** The blank itself */
const SLOT =
  '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>';

const cell = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${content}</w:p></w:tc>`;

function control(tag: string, id: number, lock = ""): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/>` +
    `<w:id w:val="${id}"/>${locked}</w:sdtPr>` +
    `<w:sdtContent>${SLOT}</w:sdtContent></w:sdt>`
  );
}

/** The blanks in the order they are drawn, which is the order `.docx-editor-empty-run` finds them */
export const EMPTY_RUN_BLANKS = {
  nameCell: 0,
  roleCell: 1,
  label: 2,
  signer: 3,
  witness: 4,
  shut: 5,
  erase: 6,
} as const;

export function emptyRunFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}"><w:body>` +
    "<w:tbl>" +
    '<w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>' +
    '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>' +
    '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
    '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>' +
    "</w:tblBorders></w:tblPr>" +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(run("Name"))}${cell(SLOT)}</w:tr>` +
    `<w:tr>${cell(run("Role"))}${cell(SLOT)}</w:tr>` +
    "</w:tbl>" +
    `<w:p>${run("Label: ")}${SLOT}</w:p>` +
    `<w:p>${run("Signer: ")}${control("SIGNER", 31)}</w:p>` +
    `<w:p>${run("Witness: ")}${control("WITNESS", 32)}</w:p>` +
    `<w:p>${run("Shut: ")}${control("SHUT", 33, "sdtContentLocked")}</w:p>` +
    `<w:p>${run("Erase: ")}${SLOT}</w:p>` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1247" w:bottom="1304" w:left="1361" w:header="567" w:footer="652" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
