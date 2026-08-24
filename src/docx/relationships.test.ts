import { describe, expect, it } from "vitest";
import { resolveTarget } from "./relationships";

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
