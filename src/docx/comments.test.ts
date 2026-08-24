// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { type Command, TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { bytesEqual, decode, makeDocx } from "../__testing__/docx";
import {
  addComment,
  addCommentReply,
  documentComments,
  removeComment,
  removeCommentReply,
  setCommentResolved,
  updateComment,
  updateCommentReply,
} from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import { commentParaId } from "./comments";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

const encoder = new TextEncoder();
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const COMMENTED_BODY =
  '<w:p><w:commentRangeStart w:id="4"/>' +
  '<w:r><w:t xml:space="preserve">Alpha</w:t></w:r>' +
  '<w:commentRangeEnd w:id="4"/>' +
  '<w:r><w:commentReference w:id="4"/></w:r>' +
  '<w:r><w:t xml:space="preserve"> beta</w:t></w:r></w:p>';

const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:comments xmlns:w="${W_NS}">` +
  '<w:comment w:id="4" w:author="Ada" w:initials="AL" w:date="2026-08-22T01:02:03Z">' +
  '<w:p><w:r><w:t xml:space="preserve">Check this</w:t></w:r></w:p>' +
  "</w:comment></w:comments>";

function contentTypes(includeComments: boolean): string {
  return (
    `<Types xmlns="${CONTENT_TYPES_NS}">` +
    '<Override PartName="/word/document.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (includeComments
      ? '<Override PartName="/word/comments.xml" ' +
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
      : "") +
    "</Types>"
  );
}

function commentsRelationship(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
    "</Relationships>"
  );
}

function threadedCommentDocument(): Uint8Array {
  const parts = unzipSync(makeCommentedDocx());
  parts["word/comments.xml"] = encoder.encode(
    `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
      '<w:comment w:id="4" w:author="Ada"><w:p w14:paraId="000000A1">' +
      '<w:r><w:t xml:space="preserve">Root note</w:t></w:r></w:p></w:comment>' +
      '<w:comment w:id="5" w:author="Grace"><w:p w14:paraId="000000A2">' +
      '<w:r><w:t xml:space="preserve">Reply note</w:t></w:r></w:p></w:comment>' +
      "</w:comments>"
  );
  parts["word/commentsExtended.xml"] = encoder.encode(
    `<w15:commentsEx xmlns:w15="${W15_NS}">` +
      '<w15:commentEx w15:paraId="000000A1" w15:done="1"/>' +
      '<w15:commentEx w15:paraId="000000A2" w15:paraIdParent="000000A1"/>' +
      '<w15:commentEx w15:paraId="DEADBEEF"/>' +
      "</w15:commentsEx>"
  );
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    commentsRelationship().replace(
      "</Relationships>",
      '<Relationship Id="rId6" Target="commentsExtended.xml" ' +
        'Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended"/>' +
        "</Relationships>"
    )
  );
  parts["[Content_Types].xml"] = encoder.encode(
    contentTypes(true).replace(
      "</Types>",
      '<Override PartName="/word/commentsExtended.xml" ' +
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' +
        "</Types>"
    )
  );
  return zipSync(parts);
}

function nestedThreadDocument(): Uint8Array {
  const parts = unzipSync(threadedCommentDocument());
  parts["word/comments.xml"] = encoder.encode(
    decode(parts["word/comments.xml"]).replace(
      "</w:comments>",
      '<w:comment w:id="6" w:author="Lin"><w:p w14:paraId="000000A3">' +
        "<w:r><w:t>Nested reply</w:t></w:r></w:p></w:comment></w:comments>"
    )
  );
  parts["word/commentsExtended.xml"] = encoder.encode(
    decode(parts["word/commentsExtended.xml"]).replace(
      "</w15:commentsEx>",
      '<w15:commentEx w15:paraId="000000A3" w15:paraIdParent="000000A2"/>' +
        "</w15:commentsEx>"
    )
  );
  return zipSync(parts);
}

function makeCommentedDocx(body = COMMENTED_BODY): Uint8Array {
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    commentsRelationship()
  );
  parts["word/comments.xml"] = encoder.encode(COMMENTS_XML);
  parts["[Content_Types].xml"] = encoder.encode(contentTypes(true));
  return zipSync(parts);
}

function makeCommentReadyDocx(body: string): Uint8Array {
  const parts = unzipSync(makeDocx(body));
  parts["[Content_Types].xml"] = encoder.encode(contentTypes(false));
  return zipSync(parts);
}

function firstTextRange(doc: PMNode): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  doc.descendants((node, pos) => {
    if (range === null && node.isText) {
      range = { from: pos, to: pos + (node.text?.length ?? 0) };
    }
    return range === null;
  });
  if (range === null) throw new Error("no text in document");
  return range;
}

function apply(state: ReturnType<typeof createEditorState>, command: Command) {
  let next = state;
  expect(
    command(state, (transaction) => {
      next = state.apply(transaction);
    })
  ).toBe(true);
  return next;
}

describe("WordprocessingML comments", () => {
  it("imports the anchor and its related comment body without replacing the paragraph", () => {
    const { doc } = importDocx(makeCommentedDocx());
    const state = createEditorState(doc);

    expect(doc.textContent).toBe("Alpha beta");
    expect(doc.child(0).type.name).toBe("paragraph");
    expect(documentComments(state)).toEqual([
      expect.objectContaining({
        id: "4",
        author: "Ada",
        initials: "AL",
        date: "2026-08-22T01:02:03Z",
        text: "Check this",
      }),
    ]);
    expect(
      documentComments(state)[0]?.to - documentComments(state)[0]?.from
    ).toBe("Alpha".length);
  });

  it("leaves all comment parts byte-identical when nothing changed", () => {
    const bytes = makeCommentedDocx();
    const before = unzipSync(bytes);
    const opened = importDocx(bytes);
    const after = unzipSync(exportDocx(opened.doc, opened.session));

    for (const path of [
      "word/document.xml",
      "word/comments.xml",
      "word/_rels/document.xml.rels",
      "[Content_Types].xml",
    ]) {
      expect(bytesEqual(after[path], before[path]), path).toBe(true);
    }
  });

  it("updates the comment body and reads it back from comments.xml", () => {
    const opened = importDocx(makeCommentedDocx());
    const state = apply(
      createEditorState(opened.doc),
      updateComment("4", "A revised note")
    );
    const output = exportDocx(state.doc, opened.session);
    const commentsXml = decode(unzipSync(output)["word/comments.xml"]);

    expect(commentsXml).toContain("A revised note");
    expect(commentsXml).not.toContain("Check this");
    expect(
      documentComments(createEditorState(importDocx(output).doc))[0]?.text
    ).toBe("A revised note");
  });

  it("removes the anchor markers, reference and comment entry together", () => {
    const opened = importDocx(makeCommentedDocx());
    const state = apply(createEditorState(opened.doc), removeComment("4"));
    const output = exportDocx(state.doc, opened.session);
    const parts = unzipSync(output);

    expect(decode(parts["word/document.xml"])).not.toContain("commentRange");
    expect(decode(parts["word/document.xml"])).not.toContain(
      "commentReference"
    );
    expect(decode(parts["word/comments.xml"])).not.toContain("w:comment ");
    expect(documentComments(createEditorState(importDocx(output).doc))).toEqual(
      []
    );
  });

  it("adds a conforming comments part, relationship and content type", () => {
    const opened = importDocx(
      makeCommentReadyDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    const range = firstTextRange(opened.doc);
    let state = createEditorState(opened.doc);
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, range.from, range.from + 5)
      )
    );
    state = apply(
      state,
      addComment({
        text: "New note",
        author: "Grace",
        initials: "GH",
        date: "2026-08-22T02:03:04Z",
      })
    );

    const output = exportDocx(state.doc, opened.session);
    const parts = unzipSync(output);
    const documentXml = decode(parts["word/document.xml"]);
    const commentsXml = decode(parts["word/comments.xml"]);

    expect(documentXml).toMatch(
      /<w:commentRangeStart w:id="0"\/>.*Alpha.*<w:commentRangeEnd w:id="0"\/>.*<w:commentReference w:id="0"\/>/
    );
    expect(commentsXml).toContain('w:id="0"');
    expect(commentsXml).toContain('w:author="Grace"');
    expect(commentsXml).toContain("New note");
    expect(decode(parts["word/_rels/document.xml.rels"])).toContain(
      `${REL_BASE}/comments`
    );
    expect(decode(parts["[Content_Types].xml"])).toContain(
      "wordprocessingml.comments+xml"
    );

    const reopened = importDocx(output);
    expect(documentComments(createEditorState(reopened.doc))[0]).toEqual(
      expect.objectContaining({ author: "Grace", text: "New note" })
    );
  });

  it("keeps the prefix used by a prefixed content-types root", () => {
    const parts = unzipSync(
      makeCommentReadyDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    parts["[Content_Types].xml"] = encoder.encode(
      decode(parts["[Content_Types].xml"])
        .replace(
          `<Types xmlns="${CONTENT_TYPES_NS}">`,
          `<ct:Types xmlns:ct="${CONTENT_TYPES_NS}">`
        )
        .replaceAll("<Override", "<ct:Override")
        .replace("</Types>", "</ct:Types>")
    );
    const opened = importDocx(zipSync(parts));
    const range = firstTextRange(opened.doc);
    let state = createEditorState(opened.doc);
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, range.from, range.from + 5)
      )
    );
    state = apply(state, addComment({ text: "New note", author: "Grace" }));

    expect(
      decode(
        unzipSync(exportDocx(state.doc, opened.session))["[Content_Types].xml"]
      )
    ).toContain('<ct:Override PartName="/word/comments.xml"');
  });

  it("does not reuse an id held only by an orphan Comments-part entry", () => {
    const opened = importDocx(
      makeCommentedDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    const range = firstTextRange(opened.doc);
    let state = createEditorState(opened.doc, {
      reservedCommentIds: opened.session.comments.byId.keys(),
    });
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, range.from, range.from + 5)
      )
    );
    state = apply(state, addComment({ text: "New note", author: "Grace" }));

    const output = unzipSync(exportDocx(state.doc, opened.session));
    const commentsXml = decode(output["word/comments.xml"]);
    expect(commentsXml).toContain('w:id="4"');
    expect(commentsXml).toContain("Check this");
    expect(commentsXml).toContain('w:id="5"');
    expect(commentsXml).toContain("New note");
  });

  it("allocates a decimal comment id without losing large integer precision", () => {
    const opened = importDocx(
      makeCommentReadyDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>'
      )
    );
    const range = firstTextRange(opened.doc);
    let state = createEditorState(opened.doc, {
      reservedCommentIds: ["9007199254740992"],
    });
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, range.from, range.from + 5)
      )
    );
    state = apply(state, addComment({ text: "New note", author: "Grace" }));

    expect(documentComments(state)[0]?.id).toBe("9007199254740993");
  });

  it("keeps generated paragraph ids nonzero and below the signed 32-bit boundary", () => {
    for (const seed of ["comment-4", "comment-5", "comment-100", "reply"]) {
      const value = Number.parseInt(commentParaId(seed), 16);
      expect(value).toBeGreaterThan(0);
      expect(value).toBeLessThan(0x80000000);
    }
  });

  it("imports replies and the resolved state from commentsExtended.xml", () => {
    const bytes = threadedCommentDocument();
    const opened = importDocx(bytes);
    const comments = documentComments(createEditorState(opened.doc));

    expect(comments[0]).toEqual(
      expect.objectContaining({
        id: "4",
        text: "Root note",
        resolved: true,
        replies: [
          expect.objectContaining({
            id: "5",
            author: "Grace",
            text: "Reply note",
          }),
        ],
      })
    );
    const before = unzipSync(bytes);
    const after = unzipSync(exportDocx(opened.doc, opened.session));
    expect(
      bytesEqual(
        after["word/commentsExtended.xml"],
        before["word/commentsExtended.xml"]
      )
    ).toBe(true);
  });

  it("resolves and reopens a thread without deleting its anchor", () => {
    const opened = importDocx(threadedCommentDocument());
    let state = createEditorState(opened.doc);
    state = apply(state, setCommentResolved("4", false));
    let output = exportDocx(state.doc, opened.session);
    expect(decode(unzipSync(output)["word/commentsExtended.xml"])).toContain(
      'w15:done="0"'
    );
    expect(decode(unzipSync(output)["word/commentsExtended.xml"])).toContain(
      'w15:paraId="DEADBEEF"'
    );
    expect(decode(unzipSync(output)["word/document.xml"])).toContain(
      "commentReference"
    );

    const reopened = importDocx(output);
    state = apply(
      createEditorState(reopened.doc, {
        reservedCommentIds: reopened.session.comments.byId.keys(),
      }),
      setCommentResolved("4", true)
    );
    output = exportDocx(state.doc, reopened.session);
    expect(decode(unzipSync(output)["word/commentsExtended.xml"])).toContain(
      'w15:done="1"'
    );
  });

  it("adds, edits and removes a reply while preserving the root comment", () => {
    const opened = importDocx(makeCommentedDocx());
    let state = createEditorState(opened.doc, {
      reservedCommentIds: opened.session.comments.byId.keys(),
    });
    state = apply(
      state,
      addCommentReply("4", {
        text: "First reply",
        author: "Grace",
        date: "2026-08-22T03:04:05Z",
      })
    );
    expect(documentComments(state)[0]?.replies[0]?.text).toBe("First reply");
    state = apply(state, updateCommentReply("4", "5", "Edited reply"));

    let output = exportDocx(state.doc, opened.session);
    let parts = unzipSync(output);
    const commentsXml = decode(parts["word/comments.xml"]);
    expect(commentsXml).toContain("Edited reply");
    expect(commentsXml).toContain(`xmlns:mc="${MC_NS}"`);
    expect(commentsXml).toContain('mc:Ignorable="w14"');
    for (const match of commentsXml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)) {
      const value = Number.parseInt(match[1], 16);
      expect(value).toBeGreaterThan(0);
      expect(value).toBeLessThan(0x80000000);
    }
    expect(decode(parts["word/commentsExtended.xml"])).toContain(
      "w15:paraIdParent="
    );
    expect(decode(parts["word/_rels/document.xml.rels"])).toContain(
      "relationships/commentsExtended"
    );
    expect(decode(parts["[Content_Types].xml"])).toContain(
      "wordprocessingml.commentsExtended+xml"
    );

    const reopened = importDocx(output);
    expect(
      documentComments(createEditorState(reopened.doc))[0]?.replies[0]?.text
    ).toBe("Edited reply");
    state = apply(
      createEditorState(reopened.doc),
      removeCommentReply("4", "5")
    );
    output = exportDocx(state.doc, reopened.session);
    parts = unzipSync(output);
    expect(decode(parts["word/comments.xml"])).toContain('w:id="4"');
    expect(decode(parts["word/comments.xml"])).not.toContain('w:id="5"');
  });

  it("deletes every descendant when a nested reply is removed", () => {
    const opened = importDocx(nestedThreadDocument());
    expect(
      documentComments(createEditorState(opened.doc))[0]?.replies
    ).toHaveLength(2);
    const state = apply(
      createEditorState(opened.doc),
      removeCommentReply("4", "5")
    );
    const parts = unzipSync(exportDocx(state.doc, opened.session));
    const commentsXml = decode(parts["word/comments.xml"]);
    const extendedXml = decode(parts["word/commentsExtended.xml"]);

    expect(commentsXml).not.toContain('w:id="5"');
    expect(commentsXml).not.toContain('w:id="6"');
    expect(extendedXml).not.toContain('w15:paraId="000000A2"');
    expect(extendedXml).not.toContain('w15:paraId="000000A3"');
  });

  it("stops at a cycle in imported reply relationships", () => {
    const parts = unzipSync(threadedCommentDocument());
    parts["word/commentsExtended.xml"] = encoder.encode(
      `<w15:commentsEx xmlns:w15="${W15_NS}">` +
        '<w15:commentEx w15:paraId="000000A1" w15:paraIdParent="000000A2"/>' +
        '<w15:commentEx w15:paraId="000000A2" w15:paraIdParent="000000A1"/>' +
        "</w15:commentsEx>"
    );
    const opened = importDocx(zipSync(parts));

    expect(documentComments(createEditorState(opened.doc))[0]?.replies).toEqual(
      [expect.objectContaining({ id: "5", text: "Reply note" })]
    );
  });

  it("reopens a resolved thread when a reply is added", () => {
    const opened = importDocx(threadedCommentDocument());
    const state = apply(
      createEditorState(opened.doc, {
        reservedCommentIds: opened.session.comments.byId.keys(),
      }),
      addCommentReply("4", {
        text: "Reopened",
        author: "Grace",
        date: "2026-08-22T04:05:06Z",
      })
    );
    expect(documentComments(state)[0]?.resolved).toBe(false);
    expect(
      decode(
        unzipSync(exportDocx(state.doc, opened.session))[
          "word/commentsExtended.xml"
        ]
      )
    ).toContain('w15:done="0"');
  });

  it("uses the reference position for a comment with no explicit range", () => {
    const bytes = makeCommentedDocx(
      '<w:p><w:r><w:t xml:space="preserve">Alpha</w:t></w:r>' +
        '<w:r><w:commentReference w:id="4"/></w:r></w:p>'
    );
    const comment = documentComments(
      createEditorState(importDocx(bytes).doc)
    )[0];

    expect(comment?.from).toBe(comment?.referencePos);
    expect(comment?.to).toBe(comment?.referencePos);
  });

  it.each([
    ["start", '<w:commentRangeStart w:id="4"/>'],
    ["end", '<w:commentRangeEnd w:id="4"/>'],
  ])("uses an unmatched %s marker as the point anchor", (_kind, marker) => {
    const bytes = makeCommentedDocx(
      `<w:p><w:r><w:t xml:space="preserve">Alpha</w:t></w:r>${marker}` +
        '<w:r><w:t xml:space="preserve"> beta</w:t></w:r>' +
        '<w:r><w:commentReference w:id="4"/></w:r></w:p>'
    );
    const comment = documentComments(
      createEditorState(importDocx(bytes).doc)
    )[0];

    expect(comment?.from).toBe(comment?.to);
    expect(comment?.from).not.toBe(comment?.referencePos);
  });
});
