import { describe, expect, it } from "vitest";
import { docxSchema } from "./index";
import {
  asStoryKey,
  STORIES_ATTR,
  sameStory,
  storiesOf,
  storyKey,
  storyNodeOf,
  withoutCommentStories,
} from "./stories";

const paragraph = (text: string, attrs: Record<string, unknown> = {}) =>
  docxSchema.nodes.paragraph.create(attrs, docxSchema.text(text));

const story = (text: string, attrs: Record<string, unknown> = {}) =>
  docxSchema.nodes.doc.create(null, [paragraph(text, attrs)]).toJSON();

const documentHolding = (stories: Record<string, unknown>) =>
  docxSchema.nodes.doc.create({ [STORIES_ATTR]: stories }, [paragraph("Body")]);

describe("naming a story", () => {
  it("storyKey joins kind and id", () => {
    expect(storyKey("comment", "4")).toBe("comment:4");
    expect(storyKey("header", "word/header1.xml")).toBe(
      "header:word/header1.xml"
    );
  });

  it("reads a key back, an id carrying colons of its own included", () => {
    expect(asStoryKey("footnote:2")).toBe("footnote:2");
    expect(asStoryKey("header:word/header1.xml")).toBe(
      "header:word/header1.xml"
    );
    for (const text of ["body", "comment", ":4", "note:1", ""]) {
      expect(asStoryKey(text)).toBeNull();
    }
  });
});

describe("what a document holds", () => {
  it("passes over a key naming no kind of story and a value that is no document", () => {
    const doc = documentHolding({
      "comment:4": story("Note"),
      "note:1": story("Elsewhere"),
      "comment:5": "not a document",
    });

    expect(Object.keys(storiesOf(doc))).toEqual(["comment:4"]);
  });

  it("builds the same node twice for the same JSON value", () => {
    const doc = documentHolding({ "comment:4": story("Note") });

    expect(storyNodeOf(doc, "comment:4")).toBe(storyNodeOf(doc, "comment:4"));
    expect(storyNodeOf(doc, "comment:9")).toBeNull();
  });
});

describe("withoutCommentStories", () => {
  it("takes out comment and reply stories alone", () => {
    const doc = documentHolding({
      "comment:4": story("Note"),
      "comment:5": story("Reply"),
      "footnote:1": story("Under the line"),
    });

    expect(Object.keys(storiesOf(withoutCommentStories(doc)))).toEqual([
      "footnote:1",
    ]);
  });

  it("leaves footnote and header stories, every other doc attr and the content alone", () => {
    const doc = docxSchema.nodes.doc.create(
      {
        sectPr: "<w:sectPr/>",
        [STORIES_ATTR]: {
          "comment:4": story("Note"),
          "footnote:1": story("Under the line"),
          "header:word/header1.xml": story("Above the page"),
        },
      },
      [paragraph("Body")]
    );

    const stripped = withoutCommentStories(doc);

    expect(Object.keys(storiesOf(stripped)).sort()).toEqual([
      "footnote:1",
      "header:word/header1.xml",
    ]);
    expect(stripped.attrs.sectPr).toBe("<w:sectPr/>");
    expect(stripped.content.eq(doc.content)).toBe(true);
  });

  it("hands back the document itself where it holds no comment story", () => {
    const doc = documentHolding({ "footnote:1": story("Under the line") });

    expect(withoutCommentStories(doc)).toBe(doc);
  });
});

describe("comparing two stories", () => {
  it("reads past the block key, which names the document a story was read from", () => {
    const one = docxSchema.nodes.doc.create(null, [
      paragraph("Note", { srcId: "d1:comment:4:0" }),
    ]);
    const other = docxSchema.nodes.doc.create(null, [
      paragraph("Note", { srcId: "d2:comment:4:0" }),
    ]);

    expect(one.eq(other)).toBe(false);
    expect(sameStory(one, other)).toBe(true);
  });

  it("reads a rewritten body and a body that gained formatting as changed", () => {
    const plain = docxSchema.nodes.doc.create(null, [paragraph("Note")]);
    const styled = docxSchema.nodes.doc.create(null, [
      paragraph("Note", {
        pPr: '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>',
      }),
    ]);

    expect(sameStory(plain, styled)).toBe(false);
    expect(
      sameStory(plain, docxSchema.nodes.doc.create(null, [paragraph("Other")]))
    ).toBe(false);
  });

  it("answers a body the document holds none of against another as unequal", () => {
    const held = docxSchema.nodes.doc.create(null, [paragraph("Note")]);

    expect(sameStory(null, null)).toBe(true);
    expect(sameStory(held, null)).toBe(false);
    expect(sameStory(null, held)).toBe(false);
  });
});
