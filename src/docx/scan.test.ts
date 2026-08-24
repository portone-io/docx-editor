// @vitest-environment node
import { describe, expect, it } from "vitest";
import { DocxImportError } from "../ooxml/errors";
import { scanBody } from "./scan";

function doc(body: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}</w:body></w:document>`
  );
}

function reconstruct(source: string): string {
  const scan = scanBody(source);
  return scan.prefix + scan.blocks.map((b) => b.xml).join("") + scan.suffix;
}

describe("scanBody", () => {
  it("returns the fragments of the body's direct children exactly as they were written", () => {
    const source = doc('<w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr w:x="1"/>');
    const scan = scanBody(source);
    expect(scan.blocks.map((b) => b.name)).toEqual(["w:p", "w:sectPr"]);
    expect(scan.blocks[0].xml).toBe("<w:p><w:r><w:t>a</w:t></w:r></w:p>");
    expect(scan.blocks[1].xml).toBe('<w:sectPr w:x="1"/>');
    expect(reconstruct(source)).toBe(source);
  });

  it("finds the boundary exactly even when a block swallows a descendant of the same name", () => {
    const source = doc(
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"
    );
    const scan = scanBody(source);
    expect(scan.blocks.map((b) => b.name)).toEqual(["w:tbl", "w:p"]);
    expect(reconstruct(source)).toBe(source);
  });

  it("does not read a > or a quote inside an attribute value as the end of the tag", () => {
    const source = doc(
      '<w:p w:a="x>y" w:b="c\'d"><w:r><w:t>t</w:t></w:r></w:p>'
    );
    const scan = scanBody(source);
    expect(scan.blocks).toHaveLength(1);
    expect(reconstruct(source)).toBe(source);
  });

  it("does not lose the whitespace and comments between blocks", () => {
    const source = doc("\n  <!-- note -->\n  <w:p/>\n  <w:sectPr/>\n");
    const scan = scanBody(source);
    expect(scan.blocks.map((b) => b.name)).toEqual(["w:p", "w:sectPr"]);
    expect(reconstruct(source)).toBe(source);
  });

  it("throws when there is no body", () => {
    expect(() => scanBody("<a><b/></a>")).toThrow(DocxImportError);
  });

  it("throws when the tags do not pair up", () => {
    expect(() => scanBody(doc("<w:p><w:r></w:p></w:r>"))).toThrow(
      DocxImportError
    );
  });
});
