// @vitest-environment jsdom
/**
 * Every fragment the schema draws has to be one the schema reads back.
 *
 * `prosemirror-view` re-reads the live DOM through these very rules after an IME composition and
 * after a browser-native edit (`parseBetween`), so a value `toDOM` writes that `getAttrs` turns
 * down is a node that disappears under someone's cursor mid-sentence. Held against whole
 * documents rather than a hand-built tree, so a fragment only a real file produces is covered too.
 */

import { unzipSync, zipSync } from "fflate";
import {
  DOMSerializer,
  DOMParser as PMDOMParser,
  type Node as PMNode,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  fixtureNames,
  makeDocx,
  makeLinkedDocx,
  readFixture,
} from "../__testing__/docx";
import { exportDocx } from "../docx/exportDocx";
import { importDocx } from "../docx/importDocx";
import { R_NS, W_NS } from "../ooxml/xml";
import { docxSchema } from "./index";

const serializer = DOMSerializer.fromSchema(docxSchema);
const parser = PMDOMParser.fromSchema(docxSchema);

/** Every node a tree holds with the marks it wears, so a refusal names what went missing */
function shapes(node: PMNode): string[] {
  const found: string[] = [];
  node.descendants((child) => {
    const marks = child.marks.map((mark) => mark.type.name).join(",");
    found.push(marks === "" ? child.type.name : `${child.type.name}[${marks}]`);
    return true;
  });
  return found;
}

describe("what the schema draws, the schema reads back", () => {
  it.each(["wx", ""])(
    "keeps formatting imported under the %j WordprocessingML prefix",
    (prefix) => {
      const parts = unzipSync(
        makeDocx(
          '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>formatted</w:t></w:r></w:p>'
        )
      );
      const path = "word/document.xml";
      const xml = new TextDecoder().decode(parts[path]);
      // Default namespaces apply to elements only; WordprocessingML attributes retain w:.
      const declaration = prefix === "" ? "xmlns=" : `xmlns:${prefix}=`;
      const qualified = prefix === "" ? "" : `${prefix}:`;
      const renamed = xml
        .replace("xmlns:w=", declaration)
        .replaceAll("<w:", `<${qualified}`)
        .replaceAll("</w:", `</${qualified}`);
      parts[path] = new TextEncoder().encode(
        prefix === ""
          ? renamed.replace("<document ", `<document xmlns:w="${W_NS}" `)
          : renamed.replaceAll(" w:", ` ${qualified}`)
      );
      const bytes = zipSync(parts);
      const { doc, session } = importDocx(bytes);
      expect(
        doc.firstChild?.firstChild?.marks.some(
          (mark) => mark.attrs.format?.bold
        )
      ).toBe(true);
      const host = document.createElement("div");
      host.appendChild(serializer.serializeFragment(doc.content));
      const reparsed = parser.parse(host, { preserveWhitespace: true });
      expect(reparsed.eq(doc)).toBe(true);
      expect(
        bytesEqual(unzipSync(exportDocx(reparsed, session))[path], parts[path])
      ).toBe(true);
    }
  );

  it("drops a control that rebinds relationships while preserving the link it wrapped", () => {
    const bytes = makeLinkedDocx(
      '<w:p><w:sdt><w:sdtPr/><w:sdtContent><w:hyperlink r:id="rId7"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:sdtContent></w:sdt></w:p>',
      { rId7: "https://example.com" }
    );
    const { doc, session } = importDocx(bytes);
    const host = document.createElement("div");
    host.appendChild(serializer.serializeFragment(doc.content));
    const control = host.querySelector("[data-sdt-prefix]");
    expect(control).not.toBeNull();
    control?.setAttribute(
      "data-sdt-prefix",
      `<w:sdt xmlns:r="urn:evil"><w:sdtPr xmlns:r="${R_NS}"/>`
    );
    const reparsed = parser.parse(host, { preserveWhitespace: true });
    expect(
      reparsed.firstChild?.firstChild?.marks.some(
        (mark) => mark.type.name === "sdt"
      )
    ).toBe(false);
    const reopened = importDocx(exportDocx(reparsed, session)).doc;
    expect(reopened.textContent).toBe("link");
    expect(
      reopened.firstChild?.firstChild?.marks.find(
        (mark) => mark.type.name === "link"
      )?.attrs.href
    ).toBe("https://example.com");
  });

  it("has fixtures to read", () => {
    expect(fixtureNames.length).toBeGreaterThan(0);
  });

  it.each(fixtureNames)("%s round trips through the DOM", (name) => {
    const { doc } = importDocx(readFixture(name));
    const host = document.createElement("div");
    host.appendChild(serializer.serializeFragment(doc.content));

    const reparsed = parser.parse(host, { preserveWhitespace: true });

    expect(shapes(reparsed)).toEqual(shapes(doc));
    expect(reparsed.content.eq(doc.content)).toBe(true);
  });
});
