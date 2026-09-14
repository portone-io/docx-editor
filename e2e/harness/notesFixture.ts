/**
 * Builds the document the footnote placement test opens, over the demo's own package so its header,
 * footer, and paper stay: pages of text referring to a footnote near the top and to a formatted one
 * a few pages down, and to an endnote written in italics.
 */

import { unzipSync, zipSync } from "fflate";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function paragraph(text: string, after = ""): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>${after}</w:p>`;
}

const footnoteReference = (id: number) =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
const endnoteReference = (id: number) =>
  `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;

const FOOTNOTES =
  `<w:footnotes ${W_NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>' +
  '<w:r><w:t xml:space="preserve"> A footnote near the top of the document.</w:t></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>' +
  '<w:r><w:t xml:space="preserve"> Formatted with </w:t></w:r>' +
  "<w:r><w:rPr><w:b/></w:rPr><w:t>bold words</w:t></w:r></w:p>" +
  "<w:p><w:r><w:t>and a second paragraph.</w:t></w:r></w:p></w:footnote>" +
  "</w:footnotes>";

const ENDNOTES =
  `<w:endnotes ${W_NS}>` +
  '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
  '<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> An endnote in italics.</w:t></w:r></w:p></w:endnote>' +
  "</w:endnotes>";

export function notesFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const documentPart = parts["word/document.xml"];
  if (!documentPart) throw new Error("the base package holds no document");
  const original = decoder.decode(documentPart);
  const root = original.match(/<w:document[^>]*>/)?.[0];
  const sectionStart = original.lastIndexOf("<w:sectPr");
  const sectionEnd = original.lastIndexOf("</w:sectPr>");
  if (root === undefined || sectionStart < 0 || sectionEnd < sectionStart) {
    throw new Error("the base document has no root or no final section");
  }
  // The final section keeps the demo's header, footer, and paper
  const finalSection = original.slice(
    sectionStart,
    sectionEnd + "</w:sectPr>".length
  );
  const body = Array.from({ length: 150 }, (_, index) => {
    const text = `Paragraph ${index + 1} keeps the text running down the page.`;
    if (index === 4) return paragraph(text, footnoteReference(1));
    if (index === 110) return paragraph(text, footnoteReference(2));
    if (index === 140) return paragraph(text, endnoteReference(1));
    return paragraph(text);
  }).join("");
  parts["word/document.xml"] = encoder.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${root}<w:body>${body}${finalSection}</w:body></w:document>`
  );
  parts["word/footnotes.xml"] = encoder.encode(FOOTNOTES);
  parts["word/endnotes.xml"] = encoder.encode(ENDNOTES);
  return zipSync(parts);
}
