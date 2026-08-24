/** Builds a compact document with a non-default automatic tab interval. */

import { unzipSync, zipSync } from "fflate";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function tabStop(value: string, positionTwips: number): string {
  return `<w:tabs><w:tab w:val="${value}" w:pos="${positionTwips}"/></w:tabs>`;
}

function paragraph(properties: string, content: string): string {
  return `<w:p><w:pPr>${properties}</w:pPr><w:r>${content}</w:r></w:p>`;
}

function paragraphRuns(properties: string, runs: string): string {
  return `<w:p><w:pPr>${properties}</w:pPr>${runs}</w:p>`;
}

const text = (value: string) => `<w:t>${value}</w:t>`;
const TAB = "<w:tab/>";

export function tabFixture(base: Uint8Array): Uint8Array {
  const parts = unzipSync(base);
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
    paragraph("", text("A") + TAB + text("Default")) +
    paragraph(tabStop("start", 1440), text("A") + TAB + text("Start")) +
    paragraph(tabStop("center", 2880), text("A") + TAB + text("Center")) +
    paragraph(tabStop("end", 2880), text("A") + TAB + text("End")) +
    paragraph(tabStop("decimal", 2880), text("A") + TAB + text("123.45")) +
    paragraph(tabStop("num", 1440), text("A") + TAB + text("List")) +
    paragraph(
      '<w:tabs><w:tab w:val="bar" w:pos="720"/>' +
        '<w:tab w:val="start" w:pos="1440"/></w:tabs>',
      text("A") + TAB + text("Bar skipped")
    ) +
    paragraph(
      tabStop("start", 1440),
      text("A") +
        TAB +
        text("One") +
        '<w:br w:type="textWrapping"/>' +
        text("A") +
        TAB +
        text("Two")
    ) +
    '<w:p><w:hyperlink r:id="rId9"><w:r>' +
    text("https://example.com") +
    "</w:r></w:hyperlink></w:p>" +
    paragraph(
      `<w:bidi/>${tabStop("start", 1440)}`,
      text("א") + TAB + text("אבג")
    ) +
    paragraph(
      `<w:bidi/>${tabStop("end", 2880)}`,
      text("א") + TAB + text("אבג")
    ) +
    paragraph('<w:ind w:start="600"/>', text("A") + TAB + text("Indented")) +
    paragraph("", text("W".repeat(47)) + TAB + text("Wrapped")) +
    paragraphRuns(
      tabStop("decimal", 2880),
      `<w:r>${text("A")}${TAB}</w:r>` +
        '<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr>' +
        text("1,234.56") +
        "</w:r>"
    ) +
    paragraphRuns(
      tabStop("decimal", 2880),
      `<w:r>${text("A")}${TAB}</w:r>` +
        '<w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr>' +
        text("1.234,56") +
        "</w:r>"
    ) +
    paragraph(
      '<w:bidi/><w:ind w:start="600"/>',
      text("א") + TAB + text("אבג")
    ) +
    paragraph('<w:bidi/><w:ind w:left="600"/>', text("א") + TAB + text("אבג")) +
    paragraph(
      "",
      text(",") + TAB + text(",") + TAB + text(",") + TAB + text(",")
    ) +
    paragraph("", TAB + TAB + TAB + TAB) +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1247" w:bottom="1304" w:left="1361" w:header="567" w:footer="652" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";
  const settingsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="960"/></w:settings>`;
  const encode = (value: string) => new TextEncoder().encode(value);
  parts["word/document.xml"] = encode(documentXml);
  parts["word/settings.xml"] = encode(settingsXml);
  return zipSync(parts);
}
