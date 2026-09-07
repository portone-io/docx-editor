// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  decode,
  exportErrorCode,
  makeDocx,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import {
  addComment,
  setCommentResolved,
} from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import { insertImage } from "../editor/insertImage";
import { decodeUtf8, encodeUtf8, R_NS } from "../ooxml/xml";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import {
  availablePartPath,
  CONTENT_TYPES_PATH,
  contentTypeWriter,
  readPart,
  relatedPartPath,
} from "./packageParts";

const MAIN_PART = "word/document.xml";
const TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const DOCUMENT_OVERRIDE =
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>';
const COMMENTS_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const PEOPLE_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";
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

describe("contentTypeWriter", () => {
  function packageWith(types: string): Map<string, Uint8Array> {
    return new Map([[CONTENT_TYPES_PATH, encodeUtf8(types, false)]]);
  }

  function written(part: Uint8Array | null): string {
    if (part === null) throw new Error("nothing was declared");
    return decodeUtf8(part).text;
  }

  it("declares an override once even when two planners ask", () => {
    const writer = contentTypeWriter(
      packageWith(`<Types xmlns="${TYPES_NS}">${DOCUMENT_OVERRIDE}</Types>`)
    );
    writer.addOverride("word/comments.xml", COMMENTS_TYPE);
    writer.addOverride("word/comments.xml", COMMENTS_TYPE);
    expect(written(writer.part())).toBe(
      `<Types xmlns="${TYPES_NS}">` +
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/>` +
        `${DOCUMENT_OVERRIDE}</Types>`
    );
  });

  it("answers null when nothing was asked, and when everything asked for is declared already", () => {
    const parts = packageWith(
      `<Types xmlns="${TYPES_NS}"><Default Extension="png" ContentType="image/png"/>${DOCUMENT_OVERRIDE}</Types>`
    );
    expect(contentTypeWriter(parts).part()).toBeNull();

    const writer = contentTypeWriter(parts);
    writer.addDefault("png", "image/png");
    writer.addOverride("Word/Document.xml", "anything");
    expect(writer.part()).toBeNull();
  });

  it("adds a Default for a media extension the package lacks", () => {
    const writer = contentTypeWriter(
      packageWith(
        `<Types xmlns="${TYPES_NS}"><Default Extension="PNG" ContentType="image/png"/>${DOCUMENT_OVERRIDE}</Types>`
      )
    );
    writer.addDefault("png", "image/png");
    writer.addDefault("gif", "image/gif");
    expect(written(writer.part())).toBe(
      `<Types xmlns="${TYPES_NS}">` +
        '<Default Extension="gif" ContentType="image/gif"/>' +
        '<Default Extension="PNG" ContentType="image/png"/>' +
        `${DOCUMENT_OVERRIDE}</Types>`
    );
  });

  it("writes a declaration under the prefix the root carries", () => {
    const writer = contentTypeWriter(
      packageWith(`<ct:Types xmlns:ct="${TYPES_NS}"></ct:Types>`)
    );
    writer.addDefault("png", "image/png");
    writer.addOverride("word/comments.xml", COMMENTS_TYPE);
    expect(written(writer.part())).toBe(
      `<ct:Types xmlns:ct="${TYPES_NS}">` +
        '<ct:Default Extension="png" ContentType="image/png"/>' +
        `<ct:Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/>` +
        "</ct:Types>"
    );
  });

  it("refuses to declare a part in a package that has no content types part", () => {
    const writer = contentTypeWriter(new Map());
    expect(writer.part()).toBeNull();
    writer.addOverride("word/comments.xml", COMMENTS_TYPE);
    expect(exportErrorCode(() => writer.part())).toBe("missing-content-types");
  });

  function applied(state: EditorState, command: Command): EditorState {
    let next = state;
    expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
    return next;
  }

  it("writes defaults ahead of overrides when one export adds both", () => {
    const parts = unzipSync(
      makeDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    parts[CONTENT_TYPES_PATH] = encoder.encode(
      `<Types xmlns="${TYPES_NS}">${DOCUMENT_OVERRIDE}</Types>`
    );
    const opened = importDocx(zipSync(parts));
    let state = createEditorState(opened.doc);
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1))
    );
    state = applied(
      state,
      insertImage({
        src: TINY_PNG_DATA_URL,
        extent: { cx: 952500, cy: 952500 },
      })
    );
    const { from, to } = rangeOfText(state.doc, "beta");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to))
    );
    state = applied(state, addComment({ text: "Note", author: "Grace" }));

    const output = unzipSync(exportDocx(state.doc, opened.session));
    expect(decode(output[CONTENT_TYPES_PATH])).toBe(
      `<Types xmlns="${TYPES_NS}">` +
        '<Default Extension="png" ContentType="image/png"/>' +
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/>` +
        `${DOCUMENT_OVERRIDE}</Types>`
    );
  });

  /**
   * A comment settled by an author the file has yet to record adds three parts, and each writer
   * used to prepend its own declaration, so the three came out in the reverse order of writing.
   */
  it("writes several overrides in the order their parts were added", () => {
    const parts = unzipSync(
      makeDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    parts[CONTENT_TYPES_PATH] = encoder.encode(
      `<Types xmlns="${TYPES_NS}">${DOCUMENT_OVERRIDE}</Types>`
    );
    const opened = importDocx(zipSync(parts));
    let state = createEditorState(opened.doc);
    const { from, to } = rangeOfText(state.doc, "beta");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to))
    );
    state = applied(
      state,
      addComment({ text: "Note", author: "Grace", authorId: "u_grace" })
    );
    state = applied(state, setCommentResolved("0", true));

    const output = unzipSync(exportDocx(state.doc, opened.session));
    expect(decode(output[CONTENT_TYPES_PATH])).toBe(
      `<Types xmlns="${TYPES_NS}">` +
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_TYPE}"/>` +
        `<Override PartName="/word/commentsExtended.xml" ContentType="${COMMENTS_EXTENDED_TYPE}"/>` +
        `<Override PartName="/word/people.xml" ContentType="${PEOPLE_TYPE}"/>` +
        `${DOCUMENT_OVERRIDE}</Types>`
    );
  });
});
