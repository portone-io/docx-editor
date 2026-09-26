/**
 * Builds a document whose paragraphs hold inline content controls with a bracketed placeholder
 * inside, the shape a template fills in: one between words, one standing alone in its paragraph
 * with the empty run a word processor leaves after it, and one locked against deletion alone.
 */

import { unzipSync, zipSync } from "fflate";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** What each control holds when the document opens */
export const INLINE_PLACEHOLDER = "[          ]";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const EMPTY_RUN = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';

function control(tag: string, id: number, lock = ""): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/>` +
    `<w:id w:val="${id}"/>${locked}<w:text/></w:sdtPr>` +
    `<w:sdtContent>${run(INLINE_PLACEHOLDER)}</w:sdtContent></w:sdt>`
  );
}

export function inlineControlFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}"><w:body>` +
    `<w:p>${run("Starts on ")}${control("START", 11)}${run(" at noon.")}</w:p>` +
    `<w:p>${control("PERIOD", 12)}${EMPTY_RUN}</w:p>` +
    `<w:p>${run("Signed on ")}${control("SIGNED", 13, "sdtLocked")}${run(".")}</w:p>` +
    "</w:body></w:document>";
  parts["word/document.xml"] = new TextEncoder().encode(documentXml);
  return zipSync(parts);
}
