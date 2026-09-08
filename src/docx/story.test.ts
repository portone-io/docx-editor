// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { docxSchema } from "../schema";
import { storyKey } from "../schema/stories";
import { importDocx } from "./importDocx";
import type { SessionStore } from "./session";
import { type ImportedStory, storyFromText, storyOf, storyText } from "./story";

const encoder = new TextEncoder();
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const BODY =
  '<w:p><w:commentRangeStart w:id="4"/>' +
  '<w:r><w:t xml:space="preserve">Alpha</w:t></w:r>' +
  '<w:commentRangeEnd w:id="4"/>' +
  '<w:r><w:commentReference w:id="4"/></w:r></w:p>';

/** A body with a bold run, a paragraph style, a table and a second paragraph */
const RICH_BODY =
  '<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p>' +
  '<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p><w:r><w:t xml:space="preserve">Second</w:t></w:r></w:p>';

function commentedDocx(
  bodies: readonly { id: string; xml: string }[]
): Uint8Array {
  const parts = unzipSync(makeDocx(BODY));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:comments xmlns:w="${W_NS}">\n  ` +
      bodies
        .map(
          ({ id, xml }) =>
            `<w:comment w:id="${id}" w:author="Ada">${xml}</w:comment>`
        )
        .join("\n  ") +
      "\n</w:comments>"
  );
  parts["[Content_Types].xml"] = encoder.encode(
    `<Types xmlns="${CONTENT_TYPES_NS}">` +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function storyIn(session: SessionStore, id: string): ImportedStory {
  const story = session.stories.get(storyKey("comment", id));
  if (story === undefined) throw new Error(`no story for ${id}`);
  return story;
}

describe("reading a story out of a part", () => {
  it("slices every comment and its blocks verbatim from comments.xml", () => {
    const rich = `<w:p><w:r><w:t xml:space="preserve">One</w:t></w:r></w:p>`;
    const { session } = importDocx(
      commentedDocx([
        { id: "4", xml: rich },
        { id: "5", xml: RICH_BODY },
      ])
    );

    expect(storyIn(session, "4").xml).toBe(
      `<w:comment w:id="4" w:author="Ada">${rich}</w:comment>`
    );
    const second = storyIn(session, "5");
    expect(second.blocks.map((block) => block.xml).join("")).toBe(RICH_BODY);
    expect(second.partPath).toBe("word/comments.xml");
    expect(second.kind).toBe("comment");
  });

  it("builds a story with the same paragraph and table nodes the body would get", () => {
    const { session } = importDocx(
      commentedDocx([{ id: "4", xml: RICH_BODY }])
    );
    const story = storyIn(session, "4");

    expect(story.doc.children.map((block) => block.type.name)).toEqual([
      "paragraph",
      "table",
      "paragraph",
    ]);
    expect(story.doc.child(0).attrs.pPr).toBe(
      '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
    );
    expect(story.doc.child(0).child(0).marks[0].attrs.rPr).toBe(
      "<w:rPr><w:b/></w:rPr>"
    );
  });

  it("keys every block of a story into the list the session hands out", () => {
    const { session } = importDocx(
      commentedDocx([{ id: "4", xml: RICH_BODY }])
    );
    const story = storyIn(session, "4");

    expect(story.doc.child(0).attrs.srcId).toBe(
      `${session.sessionId}:comment:4:0`
    );
    expect(session.blocksOf("comment:4")).toBe(story.blocks);
    expect(session.blocksOf("comment:9")).toEqual([]);
    expect(session.blocksOf("body")).toBe(session.blocks);
  });

  it("opens a comment written as one self-closing tag with a place to type", () => {
    const parts = unzipSync(commentedDocx([{ id: "4", xml: "" }]));
    parts["word/comments.xml"] = encoder.encode(
      `<w:comments xmlns:w="${W_NS}"><w:comment w:id="4"/></w:comments>`
    );
    const { doc, session } = importDocx(zipSync(parts));

    const story = storyIn(session, "4");
    expect(story.xml).toBe('<w:comment w:id="4"/>');
    expect(story.doc.childCount).toBe(1);
    expect(storyText(storyOf(doc, "comment:4"))).toBe("");
  });
});

describe("what a story reads as", () => {
  it("joins paragraphs with newlines and keeps tabs and breaks", () => {
    const { doc } = importDocx(
      commentedDocx([
        {
          id: "4",
          xml:
            '<w:p><w:r><w:t xml:space="preserve">a</w:t><w:tab/>' +
            '<w:t xml:space="preserve">b</w:t><w:br/>' +
            '<w:t xml:space="preserve">c</w:t></w:r></w:p>' +
            '<w:p><w:r><w:t xml:space="preserve">d</w:t></w:r></w:p>',
        },
      ])
    );

    expect(storyText(storyOf(doc, "comment:4"))).toBe("a\tb\nc\nd");
  });

  it("reads the text of a story built from plain lines back as it was written", () => {
    for (const text of ["one", "one\ntwo", "  padded  ", "댓글 \u{1F600}"]) {
      expect(storyText(storyFromText(text))).toBe(text);
    }
  });

  it("answers a document holding no such story with nothing at all", () => {
    expect(storyText(null)).toBe("");
    expect(
      storyOf(
        docxSchema.nodes.doc.create(null, [
          docxSchema.nodes.paragraph.create(),
        ]),
        "comment:4"
      )
    ).toBe(null);
  });
});
