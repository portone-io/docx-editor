// @vitest-environment jsdom
/**
 * Every fragment the schema draws has to be one the schema reads back.
 *
 * `prosemirror-view` re-reads the live DOM through these very rules after an IME composition and
 * after a browser-native edit (`parseBetween`), so a value `toDOM` writes that `getAttrs` turns
 * down is a node that disappears under someone's cursor mid-sentence. Held against whole
 * documents rather than a hand-built tree, so a fragment only a real file produces is covered too.
 */
import {
  DOMSerializer,
  DOMParser as PMDOMParser,
  type Node as PMNode,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { fixtureNames, readFixture } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
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
