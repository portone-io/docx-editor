// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { docxSchema } from "../schema";
import { storyKey } from "../schema/stories";
import { NO_EXPORT_REFS } from "./exportRefs";
import { importDocx } from "./importDocx";
import { serializeStory } from "./serializeStory";
import type { SessionStore } from "./session";
import { type ImportedStory, storyFromText } from "./story";

const encoder = new TextEncoder();
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const BODY = '<w:p><w:r><w:commentReference w:id="4"/></w:r></w:p>';

const FIRST = '<w:p><w:r><w:t xml:space="preserve">One</w:t></w:r></w:p>';
/** A second paragraph a producer spelled in a way this writer would not */
const SECOND =
  '<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Two</w:t></w:r></w:p>';

function commentedDocx(): Uint8Array {
  const parts = unzipSync(makeDocx(BODY));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(
    `<w:comments xmlns:w="${W_NS}">` +
      `<w:comment w:id="4" w:author="Ada">${FIRST}${SECOND}</w:comment>` +
      "</w:comments>"
  );
  parts["[Content_Types].xml"] = encoder.encode(
    `<Types xmlns="${CONTENT_TYPES_NS}">` +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function opened(): { session: SessionStore; story: ImportedStory } {
  const { session } = importDocx(commentedDocx());
  const story = session.stories.get(storyKey("comment", "4"));
  if (story === undefined) throw new Error("no comment story");
  return { session, story };
}

/** The same story with its first paragraph saying something else */
function retyped(story: ImportedStory, text: string): PMNode {
  const first = story.doc.child(0);
  return story.doc.copy(
    story.doc.content.replaceChild(
      0,
      first.copy(Fragment.from(docxSchema.text(text)))
    )
  );
}

const CONTAINER = { open: '<w:comment w:id="9">', close: "</w:comment>" };

describe("writing a story back", () => {
  it("a story nobody edited goes out as its original bytes", () => {
    const { session, story } = opened();

    expect(
      serializeStory(story.doc, story, CONTAINER, {
        ...NO_EXPORT_REFS,
        session,
      })
    ).toBe(story.xml);
  });

  it("editing one paragraph of a story rebuilds only that paragraph", () => {
    const { session, story } = opened();

    const written = serializeStory(
      retyped(story, "Rewritten"),
      story,
      CONTAINER,
      {
        ...NO_EXPORT_REFS,
        session,
      }
    );

    expect(written).toBe(
      '<w:comment w:id="4" w:author="Ada">' +
        '<w:p><w:r><w:t xml:space="preserve">Rewritten</w:t></w:r></w:p>' +
        `${SECOND}</w:comment>`
    );
  });

  it("writes a story the package never held into the container it is given", () => {
    const { session } = opened();

    expect(
      serializeStory(storyFromText("New"), null, CONTAINER, {
        ...NO_EXPORT_REFS,
        session,
      })
    ).toBe(
      '<w:comment w:id="9"><w:p><w:r>' +
        '<w:t xml:space="preserve">New</w:t></w:r></w:p></w:comment>'
    );
  });

  it("writes every block itself where no session says which arrived as what", () => {
    const { story } = opened();

    expect(serializeStory(story.doc, null, CONTAINER, NO_EXPORT_REFS)).toBe(
      `<w:comment w:id="9">${FIRST}${SECOND}</w:comment>`
    );
  });
});
