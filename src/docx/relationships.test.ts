// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { decodeUtf8, encodeUtf8, R_NS } from "../ooxml/xml";
import {
  readRelationships,
  relationshipWriter,
  resolveTarget,
} from "./relationships";

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const COMMENTS_REL = `${R_NS}/comments`;

describe("an internal relationship target", () => {
  it.each([
    ["word/document.xml", "comments.xml", "word/comments.xml"],
    ["word/document.xml", "../comments.xml", "comments.xml"],
    ["word/document.xml", "./comments.xml", "word/comments.xml"],
    ["word/document.xml", "/custom/comments.xml", "custom/comments.xml"],
    ["word/document.xml", "notes/footnotes.xml", "word/notes/footnotes.xml"],
    ["word/document.xml", "../notes/footnotes.xml", "notes/footnotes.xml"],
    ["word/document.xml", "./footnotes.xml", "word/footnotes.xml"],
    ["word/document.xml", "/custom/footnotes.xml", "custom/footnotes.xml"],
  ])("resolves %s plus %s to %s", (part, target, expected) => {
    expect(resolveTarget(part, target)).toBe(expected);
  });
});

describe("adding to a relationships part the package already has", () => {
  /** The part with one comments relationship added to `original` */
  function added(original: string): string {
    const writer = relationshipWriter([]);
    writer.add({ type: COMMENTS_REL, target: "comments.xml" });
    const part = writer.part(encodeUtf8(original, false));
    if (part === null) throw new Error("nothing was added");
    return decodeUtf8(part).text;
  }

  it("adds to an existing part whose root carries a prefix", () => {
    const part = added(
      `<?xml version="1.0"?><pr:Relationships xmlns:pr="${RELS_NS}"></pr:Relationships>`
    );
    expect(part).toBe(
      `<?xml version="1.0"?><pr:Relationships xmlns:pr="${RELS_NS}">` +
        `<pr:Relationship Id="rId1" Type="${COMMENTS_REL}" Target="comments.xml"/>` +
        "</pr:Relationships>"
    );
    expect(
      readRelationships(new Map([["rels", encodeUtf8(part, false)]]), "rels")
    ).toEqual([
      {
        id: "rId1",
        type: COMMENTS_REL,
        target: "comments.xml",
        external: false,
      },
    ]);
  });

  it("opens a self-closing root to hold the first relationship", () => {
    expect(added(`<Relationships xmlns="${RELS_NS}"/>`)).toBe(
      `<Relationships xmlns="${RELS_NS}">` +
        `<Relationship Id="rId1" Type="${COMMENTS_REL}" Target="comments.xml"/>` +
        "</Relationships>"
    );
  });
});
