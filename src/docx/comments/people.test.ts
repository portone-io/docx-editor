// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { bytesEqual, decode, makeDocx } from "../../__testing__/docx";
import { rangeOfText } from "../../__testing__/editing";
import {
  addComment,
  addCommentReply,
  documentComments,
} from "../../editor/commands/commentCommands";
import { createEditorState } from "../../editor/createEditor";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import {
  COMMENT_AUTHOR_PROVIDER,
  PEOPLE_CONTENT_TYPE,
  PEOPLE_REL_TYPE,
} from "./constants";

const encoder = new TextEncoder();
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const COMMENTED_BODY =
  '<w:p><w:commentRangeStart w:id="4"/>' +
  run("Alpha") +
  '<w:commentRangeEnd w:id="4"/>' +
  '<w:r><w:commentReference w:id="4"/></w:r>' +
  run(" beta") +
  "</w:p>";

const COMMENTS_XML =
  `<w:comments xmlns:w="${W_NS}">` +
  '<w:comment w:id="4" w:author="Ada"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment>' +
  "</w:comments>";

const person = (author: string, providerId: string, userId: string) =>
  `<w15:person w15:author="${author}"><w15:presenceInfo w15:providerId="${providerId}" w15:userId="${userId}"/></w15:person>`;

const peopleXml = (persons: string) =>
  `<w15:people xmlns:w15="${W15_NS}">${persons}</w15:people>`;

function relationships(entries: string): string {
  return (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    entries +
    "</Relationships>"
  );
}

const COMMENTS_REL = `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>`;
const PEOPLE_REL = `<Relationship Id="rId6" Target="people.xml" Type="${PEOPLE_REL_TYPE}"/>`;

function contentTypes(overrides: string): string {
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    overrides +
    "</Types>"
  );
}

const COMMENTS_OVERRIDE =
  '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>';
const PEOPLE_OVERRIDE = `<Override PartName="/word/people.xml" ContentType="${PEOPLE_CONTENT_TYPE}"/>`;

/** A commented document whose people part, when given, records Ada as this person */
function commentedDocx(people: string | null): Uint8Array {
  const parts = unzipSync(makeDocx(COMMENTED_BODY));
  parts["word/comments.xml"] = encoder.encode(COMMENTS_XML);
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    relationships(COMMENTS_REL + (people === null ? "" : PEOPLE_REL))
  );
  parts["[Content_Types].xml"] = encoder.encode(
    contentTypes(COMMENTS_OVERRIDE + (people === null ? "" : PEOPLE_OVERRIDE))
  );
  if (people !== null) parts["word/people.xml"] = encoder.encode(people);
  return zipSync(parts);
}

/** A plain document with a content types part, so that a comment and its author can be added */
function plainDocx(): Uint8Array {
  const parts = unzipSync(makeDocx(`<w:p>${run("Alpha beta")}</w:p>`));
  parts["[Content_Types].xml"] = encoder.encode(contentTypes(""));
  return zipSync(parts);
}

function applied(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
}

function selecting(state: EditorState, text: string): EditorState {
  const { from, to } = rangeOfText(state.doc, text);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

function authorIdsOf(bytes: Uint8Array): (string | null)[] {
  return documentComments(createEditorState(importDocx(bytes).doc)).map(
    (comment) => comment.authorId
  );
}

describe("the people part", () => {
  it("reads the identity this editor recorded for an author", () => {
    const bytes = commentedDocx(
      peopleXml(person("Ada", COMMENT_AUTHOR_PROVIDER, "u_ada"))
    );
    expect(authorIdsOf(bytes)).toEqual(["u_ada"]);
  });

  it("reads no identity from another provider, or from no people part", () => {
    expect(
      authorIdsOf(commentedDocx(peopleXml(person("Ada", "AD", "S::ada@corp"))))
    ).toEqual([null]);
    expect(authorIdsOf(commentedDocx(null))).toEqual([null]);
  });

  it("takes this editor's record of a name over another provider's", () => {
    const bytes = commentedDocx(
      peopleXml(
        person("Ada", "AD", "S::ada@corp") +
          person("Ada", COMMENT_AUTHOR_PROVIDER, "u_ada")
      )
    );
    expect(authorIdsOf(bytes)).toEqual(["u_ada"]);
  });

  it("leaves the part byte-identical when nothing changed, and when a recorded author comments again", () => {
    const bytes = commentedDocx(
      peopleXml(person("Ada", COMMENT_AUTHOR_PROVIDER, "u_ada"))
    );
    const opened = importDocx(bytes);
    const untouched = unzipSync(exportDocx(opened.doc, opened.session));
    expect(
      bytesEqual(
        untouched["word/people.xml"],
        unzipSync(bytes)["word/people.xml"]
      )
    ).toBe(true);

    const state = applied(
      selecting(createEditorState(opened.doc), "beta"),
      addComment({ text: "Again", author: "Ada", authorId: "u_ada" })
    );
    const again = unzipSync(exportDocx(state.doc, opened.session));
    expect(
      bytesEqual(again["word/people.xml"], unzipSync(bytes)["word/people.xml"])
    ).toBe(true);
  });

  it("adds the part, its relationship and its content type for a new author", () => {
    const opened = importDocx(plainDocx());
    const state = applied(
      selecting(createEditorState(opened.doc), "beta"),
      addComment({ text: "Note", author: "Grace", authorId: "u_grace" })
    );
    const output = exportDocx(state.doc, opened.session);
    const parts = unzipSync(output);

    expect(decode(parts["word/people.xml"])).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        peopleXml(person("Grace", COMMENT_AUTHOR_PROVIDER, "u_grace"))
    );
    expect(decode(parts["word/_rels/document.xml.rels"])).toContain(
      `Type="${PEOPLE_REL_TYPE}" Target="people.xml"`
    );
    expect(decode(parts["[Content_Types].xml"])).toContain(PEOPLE_OVERRIDE);
    expect(authorIdsOf(output)).toEqual(["u_grace"]);
  });

  it("splices a new author into a part that is already there, keeping what it held", () => {
    const original = peopleXml(person("Ada", "AD", "S::ada@corp"));
    const opened = importDocx(commentedDocx(original));
    const state = applied(
      selecting(createEditorState(opened.doc), "beta"),
      addComment({ text: "Note", author: "Grace", authorId: "u_grace" })
    );
    const parts = unzipSync(exportDocx(state.doc, opened.session));
    expect(decode(parts["word/people.xml"])).toBe(
      peopleXml(
        person("Ada", "AD", "S::ada@corp") +
          person("Grace", COMMENT_AUTHOR_PROVIDER, "u_grace")
      )
    );
    expect(
      decode(parts["word/_rels/document.xml.rels"]).split("people.xml").length
    ).toBe(2);
  });

  it("records the author of a reply as well", () => {
    const opened = importDocx(commentedDocx(null));
    const state = applied(
      createEditorState(opened.doc),
      addCommentReply("4", { text: "Reply", author: "Lin", authorId: "u_lin" })
    );
    const output = exportDocx(state.doc, opened.session);
    expect(decode(unzipSync(output)["word/people.xml"])).toContain(
      person("Lin", COMMENT_AUTHOR_PROVIDER, "u_lin")
    );
    const reopened = documentComments(
      createEditorState(importDocx(output).doc)
    );
    expect(reopened[0]?.authorId).toBeNull();
    expect(reopened[0]?.replies[0]?.authorId).toBe("u_lin");
  });

  it("records nothing for a comment written under no identity", () => {
    const opened = importDocx(plainDocx());
    const state = applied(
      selecting(createEditorState(opened.doc), "beta"),
      addComment({ text: "Note", author: "Grace" })
    );
    const parts = unzipSync(exportDocx(state.doc, opened.session));
    expect(parts["word/people.xml"]).toBeUndefined();
  });
});
