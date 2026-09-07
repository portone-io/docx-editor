// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { encodeUtf8, R_NS } from "../ooxml/xml";
import { availablePartPath, readPart, relatedPartPath } from "./packageParts";

const MAIN_PART = "word/document.xml";
const encoder = new TextEncoder();

function rels(entries: string): Uint8Array {
  return encodeUtf8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      entries +
      "</Relationships>",
    false
  );
}

describe("relatedPartPath", () => {
  it("finds the part the main part relates under a type, skipping one outside the package", () => {
    const parts = new Map<string, Uint8Array>([
      [
        "word/_rels/document.xml.rels",
        rels(
          `<Relationship Id="rId1" Type="${R_NS}/comments" Target="http://example.com/comments.xml" TargetMode="External"/>` +
            `<Relationship Id="rId2" Type="${R_NS}/comments" Target="../shared/comments.xml"/>` +
            `<Relationship Id="rId3" Type="${R_NS}/styles" Target="styles.xml"/>`
        ),
      ],
      ["shared/comments.xml", encoder.encode("<w:comments/>")],
    ]);
    expect(relatedPartPath(parts, MAIN_PART, `${R_NS}/comments`)).toBe(
      "shared/comments.xml"
    );
    expect(relatedPartPath(parts, MAIN_PART, `${R_NS}/styles`)).toBe(
      "word/styles.xml"
    );
    expect(relatedPartPath(parts, MAIN_PART, `${R_NS}/numbering`)).toBeNull();
    expect(readPart(parts, "shared/comments.xml")).toBe("<w:comments/>");
    expect(readPart(parts, "word/styles.xml")).toBeNull();
    expect(readPart(parts, null)).toBeNull();
  });
});

describe("availablePartPath", () => {
  it("allocates comments.xml then comments2.xml", () => {
    const parts = new Map<string, Uint8Array>();
    expect(availablePartPath(parts, MAIN_PART, "comments")).toBe(
      "word/comments.xml"
    );
    parts.set("word/comments.xml", encoder.encode("<w:notComments/>"));
    expect(availablePartPath(parts, MAIN_PART, "comments")).toBe(
      "word/comments2.xml"
    );
    parts.set("word/comments2.xml", encoder.encode("<w:notComments/>"));
    expect(availablePartPath(parts, MAIN_PART, "comments")).toBe(
      "word/comments3.xml"
    );
    expect(availablePartPath(parts, "custom/main.xml", "comments")).toBe(
      "custom/comments.xml"
    );
  });
});
