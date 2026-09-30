// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { beforeAll, describe, expect, it } from "vitest";
import { bytesEqual, decode, makeDocx } from "../../__testing__/docx";
import { rangeOfText } from "../../__testing__/editing";
import {
  addComment,
  addCommentReply,
  documentComments,
  removeComment,
  removeCommentReply,
} from "../../editor/commands/commentCommands";
import {
  createEditorState,
  editorStateForSession,
} from "../../editor/createEditor";
import {
  reservedCommentDurableIds,
  reservedCommentParaIds,
} from "../../editor/editorDocument";
import { onlyCommentsChangedBy } from "../commentOnlyChange";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import { exportProblems } from "../invariants";
import {
  COMMENTS_EXTENSIBLE_CONTENT_TYPE,
  COMMENTS_EXTENSIBLE_REL_TYPE,
  COMMENTS_IDS_CONTENT_TYPE,
  COMMENTS_IDS_REL_TYPE,
} from "./constants";

// `vitest.config.ts` runs the suite in Seoul, nine hours ahead of UTC, so a wall-clock time read or
// written as an instant is nine hours out rather than hidden behind a machine running on UTC
beforeAll(() => {
  expect(new Date(Date.UTC(2026, 8, 30)).getTimezoneOffset()).toBe(-540);
});

const encoder = new TextEncoder();
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const COMMENTS_PATH = "word/comments.xml";
const IDS_PATH = "word/commentsIds.xml";
const EXTENSIBLE_PATH = "word/commentsExtensible.xml";
const RELS_PATH = "word/_rels/document.xml.rels";
const TYPES_PATH = "[Content_Types].xml";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const anchored = (id: string, text: string) =>
  `<w:commentRangeStart w:id="${id}"/>${run(text)}` +
  `<w:commentRangeEnd w:id="${id}"/>` +
  `<w:r><w:commentReference w:id="${id}"/></w:r>`;

const BODY =
  `<w:p>${anchored("1", "Alpha")}${run(" beta ")}${anchored("2", "gamma")}` +
  `${run(" delta")}</w:p>`;

const entry = (
  id: string,
  author: string,
  date: string,
  paraId: string,
  text: string
) =>
  `<w:comment w:id="${id}" w:author="${author}" w:date="${date}" w:initials="${author.slice(0, 2)}">` +
  `<w:p w14:paraId="${paraId}" w14:textId="77777777">${run(text)}</w:p></w:comment>`;

/**
 * What Word 365 writes for comments made seven hours behind UTC: `w:date` holds the author's
 * wall clock and `dateUtc` the instant. The second comment has a durable id and no date, the way a
 * Word that predates the extensible part leaves one.
 */
const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14">` +
  entry("1", "Ada", "2024-04-08T10:32:00Z", "1A2B3C4D", "First") +
  entry("2", "Ada", "2024-04-08T10:33:00Z", "2B3C4D5E", "Second") +
  entry("3", "Lin", "2024-04-08T11:00:00Z", "3C4D5E6F", "Reply") +
  "</w:comments>";

const EXTENDED_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w15:commentsEx xmlns:w15="${W15_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w15">` +
  '<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/>' +
  '<w15:commentEx w15:paraId="2B3C4D5E" w15:done="0"/>' +
  '<w15:commentEx w15:paraId="3C4D5E6F" w15:paraIdParent="1A2B3C4D" w15:done="0"/>' +
  "</w15:commentsEx>";

const idEntry = (paraId: string, durableId: string) =>
  `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/>`;

const dateEntry = (durableId: string, dateUtc: string) =>
  `<w16cex:commentExtensible w16cex:durableId="${durableId}" w16cex:dateUtc="${dateUtc}"/>`;

const IDS_ROOT = `<w16cid:commentsIds xmlns:mc="${MC_NS}" xmlns:w16cid="${W16CID_NS}" mc:Ignorable="w16cid">`;
const IDS_ENTRIES =
  idEntry("1A2B3C4D", "12EC8C19") +
  idEntry("2B3C4D5E", "483A933D") +
  idEntry("3C4D5E6F", "5A5A5A5A");

const idsXml = (entries: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `${IDS_ROOT}${entries}</w16cid:commentsIds>`;

const EXTENSIBLE_ROOT = `<w16cex:commentsExtensible xmlns:mc="${MC_NS}" xmlns:w16cex="${W16CEX_NS}" mc:Ignorable="w16cex">`;
const EXTENSIBLE_ENTRIES =
  dateEntry("12EC8C19", "2024-04-08T17:32:00Z") +
  dateEntry("5A5A5A5A", "2024-04-08T18:00:00Z");

const extensibleXml = (entries: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `${EXTENSIBLE_ROOT}${entries}</w16cex:commentsExtensible>`;

const override = (name: string, type: string) =>
  `<Override PartName="/word/${name}" ContentType="${type}"/>`;

function wordDocx(
  ids = idsXml(IDS_ENTRIES),
  extensible = extensibleXml(EXTENSIBLE_ENTRIES)
): Uint8Array {
  const parts = unzipSync(makeDocx(BODY));
  parts[COMMENTS_PATH] = encoder.encode(COMMENTS_XML);
  parts["word/commentsExtended.xml"] = encoder.encode(EXTENDED_XML);
  parts[IDS_PATH] = encoder.encode(ids);
  parts[EXTENSIBLE_PATH] = encoder.encode(extensible);
  parts[RELS_PATH] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      '<Relationship Id="rId6" Target="commentsExtended.xml" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended"/>' +
      `<Relationship Id="rId7" Target="commentsIds.xml" Type="${COMMENTS_IDS_REL_TYPE}"/>` +
      `<Relationship Id="rId8" Target="commentsExtensible.xml" Type="${COMMENTS_EXTENSIBLE_REL_TYPE}"/>` +
      "</Relationships>"
  );
  parts[TYPES_PATH] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      override(
        "document.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
      ) +
      override(
        "comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
      ) +
      override(
        "commentsExtended.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
      ) +
      override("commentsIds.xml", COMMENTS_IDS_CONTENT_TYPE) +
      override("commentsExtensible.xml", COMMENTS_EXTENSIBLE_CONTENT_TYPE) +
      "</Types>"
  );
  return zipSync(parts);
}

/** A document with no comment and no comment part, whose content types part takes new ones */
function plainDocx(): Uint8Array {
  const parts = unzipSync(makeDocx(`<w:p>${run("Alpha beta")}</w:p>`));
  parts[TYPES_PATH] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      override(
        "document.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
      ) +
      "</Types>"
  );
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

/** The package after `command`, run over the document `bytes` opens as */
function edited(
  bytes: Uint8Array,
  command: (state: EditorState) => EditorState
): Uint8Array {
  const opened = importDocx(bytes);
  const state = command(editorStateForSession(opened));
  expect(exportProblems(state.doc, opened.session)).toEqual([]);
  return exportDocx(state.doc, opened.session);
}

const partText = (bytes: Uint8Array, path: string): string =>
  decode(unzipSync(bytes)[path]);

function repacked(
  bytes: Uint8Array,
  path: string,
  rewrite: (text: string) => string
): Uint8Array {
  const parts = unzipSync(bytes);
  parts[path] = encoder.encode(rewrite(decode(parts[path])));
  return zipSync(parts);
}

function datesOf(bytes: Uint8Array): Record<string, string | null> {
  const comments = documentComments(createEditorState(importDocx(bytes).doc));
  return Object.fromEntries(
    comments.flatMap((comment) => [
      [comment.id, comment.date],
      ...comment.replies.map((reply) => [reply.id, reply.date]),
    ])
  );
}

/** The new comment the reopened file holds under `id`, its thread key and its durable id */
function writtenIds(
  bytes: Uint8Array,
  id: string
): { paraId: string; durableId: string } {
  const { session } = importDocx(bytes);
  const comment = session.comments.byId.get(id);
  if (comment?.paraId == null || comment.durableId === null) {
    throw new Error(`comment ${id} carries no durable id`);
  }
  return { paraId: comment.paraId, durableId: comment.durableId };
}

const WRITTEN_AT = "2026-09-30T06:00:00.000Z";

const commentedByGrace = (bytes: Uint8Array) =>
  edited(bytes, (state) =>
    applied(
      selecting(state, "delta"),
      addComment({
        text: "Mine",
        author: "Ada",
        authorId: "u_grace",
        date: WRITTEN_AT,
      })
    )
  );

describe("reading when a comment was written", () => {
  it("takes the instant from dateUtc and the wall clock alone from w:date", () => {
    expect(datesOf(wordDocx())).toEqual({
      "1": "2024-04-08T17:32:00.000Z",
      // No date reached through the durable id, so w:date's digits are read on Seoul's clock
      "2": "2024-04-08T01:33:00.000Z",
      "3": "2024-04-08T18:00:00.000Z",
    });
  });

  it("reads w:date alone where the file has neither part", () => {
    const parts = unzipSync(wordDocx());
    delete parts[IDS_PATH];
    delete parts[EXTENSIBLE_PATH];
    expect(datesOf(zipSync(parts))).toEqual({
      "1": "2024-04-08T01:32:00.000Z",
      "2": "2024-04-08T01:33:00.000Z",
      "3": "2024-04-08T02:00:00.000Z",
    });
  });

  it("follows a durable id only through the thread key the ids part names it by", () => {
    const rekeyed = wordDocx(
      idsXml(
        idEntry("0000FFFF", "12EC8C19") +
          idEntry("2B3C4D5E", "483A933D") +
          idEntry("3C4D5E6F", "5A5A5A5A")
      )
    );
    expect(datesOf(rekeyed)["1"]).toBe("2024-04-08T01:32:00.000Z");
  });

  it("writes an untouched file back byte for byte", () => {
    const bytes = wordDocx();
    const { doc, session } = importDocx(bytes);
    const output = unzipSync(exportDocx(doc, session));
    const input = unzipSync(bytes);
    for (const path of Object.keys(input)) {
      expect(bytesEqual(output[path], input[path]), path).toBe(true);
    }
  });
});

describe("writing a comment", () => {
  it("records the wall clock in w:date and the instant in the ids and extensible parts, leaving every entry that arrived as it was", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const { paraId, durableId } = writtenIds(output, "4");

    expect(partText(output, COMMENTS_PATH)).toContain(
      '<w:comment w:id="4" w:author="Ada" w:date="2026-09-30T15:00:00Z">'
    );
    expect(partText(output, COMMENTS_PATH)).toContain(`w14:paraId="${paraId}"`);
    expect(partText(output, IDS_PATH)).toBe(
      idsXml(IDS_ENTRIES + idEntry(paraId, durableId))
    );
    expect(partText(output, EXTENSIBLE_PATH)).toBe(
      extensibleXml(
        EXTENSIBLE_ENTRIES + dateEntry(durableId, "2026-09-30T06:00:00Z")
      )
    );
    expect(["12EC8C19", "483A933D", "5A5A5A5A"]).not.toContain(durableId);
    expect(datesOf(output)).toEqual({
      ...datesOf(bytes),
      "4": WRITTEN_AT,
    });
  });

  it("reads back the instant it was given, before and after the file is written", () => {
    const { doc } = importDocx(plainDocx());
    const state = applied(
      selecting(createEditorState(doc), "beta"),
      addComment({ text: "Note", author: "Ada", date: WRITTEN_AT })
    );
    expect(documentComments(state)[0]?.date).toBe(WRITTEN_AT);
    expect(datesOf(edited(plainDocx(), () => state))["0"]).toBe(WRITTEN_AT);
  });

  it("creates the ids and extensible parts, related and declared, in a file without them", () => {
    const output = edited(plainDocx(), (state) =>
      applied(
        selecting(state, "beta"),
        addComment({ text: "Note", author: "Ada", date: WRITTEN_AT })
      )
    );
    const { paraId, durableId } = writtenIds(output, "0");

    expect(partText(output, IDS_PATH)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<w16cid:commentsIds xmlns:w16cid="${W16CID_NS}">` +
        idEntry(paraId, durableId) +
        "</w16cid:commentsIds>"
    );
    expect(partText(output, EXTENSIBLE_PATH)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX_NS}">` +
        dateEntry(durableId, "2026-09-30T06:00:00Z") +
        "</w16cex:commentsExtensible>"
    );
    expect(partText(output, RELS_PATH)).toContain(
      `Type="${COMMENTS_IDS_REL_TYPE}" Target="commentsIds.xml"`
    );
    expect(partText(output, RELS_PATH)).toContain(
      `Type="${COMMENTS_EXTENSIBLE_REL_TYPE}" Target="commentsExtensible.xml"`
    );
    expect(partText(output, TYPES_PATH)).toContain(
      override("commentsIds.xml", COMMENTS_IDS_CONTENT_TYPE)
    );
    expect(partText(output, TYPES_PATH)).toContain(
      override("commentsExtensible.xml", COMMENTS_EXTENSIBLE_CONTENT_TYPE)
    );
  });

  it("dates a reply the same way, leaving the comment it answers as it was", () => {
    const bytes = wordDocx();
    const output = edited(bytes, (state) =>
      applied(
        state,
        addCommentReply("2", {
          text: "Answer",
          author: "Lin",
          date: "2026-09-30T20:30:00Z",
        })
      )
    );
    const { paraId, durableId } = writtenIds(output, "4");

    expect(partText(output, COMMENTS_PATH)).toContain(
      '<w:comment w:id="4" w:author="Lin" w:date="2026-10-01T05:30:00Z">'
    );
    expect(partText(output, IDS_PATH)).toBe(
      idsXml(IDS_ENTRIES + idEntry(paraId, durableId))
    );
    expect(partText(output, EXTENSIBLE_PATH)).toBe(
      extensibleXml(
        EXTENSIBLE_ENTRIES + dateEntry(durableId, "2026-09-30T20:30:00Z")
      )
    );
    expect(datesOf(output)["4"]).toBe("2026-09-30T20:30:00.000Z");
    expect(onlyCommentsChangedBy(bytes, output, "u_lin")).toEqual({
      ok: true,
    });
  });

  it("puts the new date ahead of an extension list closing the extensible part", () => {
    const extension =
      '<w16cex:extLst><w16:ext xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" w16:uri="{00000000-0000-0000-0000-000000000000}"/></w16cex:extLst>';
    const output = commentedByGrace(
      wordDocx(undefined, extensibleXml(EXTENSIBLE_ENTRIES + extension))
    );
    const { durableId } = writtenIds(output, "4");
    expect(partText(output, EXTENSIBLE_PATH)).toBe(
      extensibleXml(
        EXTENSIBLE_ENTRIES +
          dateEntry(durableId, "2026-09-30T06:00:00Z") +
          extension
      )
    );
  });

  it("gives a comment a durable id no entry that arrived answers to", () => {
    const first = commentedByGrace(wordDocx());
    const { durableId } = writtenIds(first, "4");
    // An entry left in the extensible part under the id the writer would otherwise take
    const orphaned = wordDocx(
      undefined,
      extensibleXml(
        EXTENSIBLE_ENTRIES + dateEntry(durableId, "2001-01-01T00:00:00Z")
      )
    );
    const output = commentedByGrace(orphaned);

    expect(writtenIds(output, "4").durableId).not.toBe(durableId);
    expect(datesOf(output)["4"]).toBe(WRITTEN_AT);
    expect(partText(output, EXTENSIBLE_PATH)).toContain(
      dateEntry(durableId, "2001-01-01T00:00:00Z")
    );
  });

  it("reserves every thread key and durable id the ids and extensible parts name, entries naming no comment among them", () => {
    const state = editorStateForSession(
      importDocx(
        wordDocx(
          idsXml(IDS_ENTRIES + idEntry("0BADF00D", "0C0FFEE0")),
          extensibleXml(
            EXTENSIBLE_ENTRIES + dateEntry("0DDBA11A", "2001-01-01T00:00:00Z")
          )
        )
      )
    );
    expect(reservedCommentParaIds(state).has("0BADF00D")).toBe(true);
    expect([...reservedCommentDurableIds(state)].sort()).toEqual([
      "0C0FFEE0",
      "0DDBA11A",
      "12EC8C19",
      "483A933D",
      "5A5A5A5A",
    ]);
  });

  it("writes w:date alone for a date that names no instant", () => {
    const output = edited(plainDocx(), (state) =>
      applied(
        selecting(state, "beta"),
        addComment({ text: "Note", author: "Ada", date: "someday" })
      )
    );
    expect(partText(output, COMMENTS_PATH)).toContain('w:date="someday"');
    expect(unzipSync(output)[IDS_PATH]).toBeUndefined();
    expect(unzipSync(output)[EXTENSIBLE_PATH]).toBeUndefined();
    expect(datesOf(output)["0"]).toBe("someday");
  });

  it("writes w:date alone into a package with no content types part to declare the new parts in", () => {
    const parts = unzipSync(wordDocx());
    delete parts[IDS_PATH];
    delete parts[EXTENSIBLE_PATH];
    delete parts[TYPES_PATH];
    const output = commentedByGrace(zipSync(parts));
    expect(unzipSync(output)[IDS_PATH]).toBeUndefined();
    expect(unzipSync(output)[EXTENSIBLE_PATH]).toBeUndefined();
    // The author's wall clock, read back on the same clock
    expect(datesOf(output)["4"]).toBe(WRITTEN_AT);
  });
});

describe("deleting a comment", () => {
  it("takes out the entries of the comment and its reply, and nothing else", () => {
    const bytes = wordDocx();
    const output = edited(bytes, (state) => applied(state, removeComment("1")));

    expect(partText(output, IDS_PATH)).toBe(
      idsXml(idEntry("2B3C4D5E", "483A933D"))
    );
    expect(partText(output, EXTENSIBLE_PATH)).toBe(extensibleXml(""));
    expect(onlyCommentsChangedBy(bytes, output, "u_ada")).toEqual({
      ok: true,
    });
  });

  it("takes out a reply's entries alone", () => {
    const bytes = wordDocx();
    const output = edited(bytes, (state) =>
      applied(state, removeCommentReply("1", "3"))
    );

    expect(partText(output, IDS_PATH)).toBe(
      idsXml(idEntry("1A2B3C4D", "12EC8C19") + idEntry("2B3C4D5E", "483A933D"))
    );
    expect(partText(output, EXTENSIBLE_PATH)).toBe(
      extensibleXml(dateEntry("12EC8C19", "2024-04-08T17:32:00Z"))
    );
    expect(datesOf(output)).toEqual({
      "1": "2024-04-08T17:32:00.000Z",
      "2": "2024-04-08T01:33:00.000Z",
    });
  });

  it("leaves the layout between the entries that stay where it stood", () => {
    const laidOut = wordDocx(
      idsXml(
        `\n  ${idEntry("1A2B3C4D", "12EC8C19")}\n  ${idEntry("2B3C4D5E", "483A933D")}` +
          `\n  ${idEntry("3C4D5E6F", "5A5A5A5A")}\n`
      )
    );
    const output = edited(laidOut, (state) =>
      applied(state, removeComment("2"))
    );
    expect(partText(output, IDS_PATH)).toBe(
      idsXml(
        `\n  ${idEntry("1A2B3C4D", "12EC8C19")}\n  ` +
          `\n  ${idEntry("3C4D5E6F", "5A5A5A5A")}\n`
      )
    );
    expect(onlyCommentsChangedBy(laidOut, output, "u_ada")).toEqual({
      ok: true,
    });
  });
});

describe("judging a returned file's durable ids and dates", () => {
  const rejected = (part: string) => ({
    ok: false,
    reason: "comment-markup-rejected",
    part,
  });

  it("takes a comment the writer dated", () => {
    const bytes = wordDocx();
    expect(
      onlyCommentsChangedBy(bytes, commentedByGrace(bytes), "u_grace")
    ).toEqual({
      ok: true,
    });
    const plain = plainDocx();
    const first = edited(plain, (state) =>
      applied(
        selecting(state, "beta"),
        addComment({ text: "Note", author: "Ada", authorId: "u_ada" })
      )
    );
    expect(onlyCommentsChangedBy(plain, first, "u_ada")).toEqual({ ok: true });
  });

  it("refuses a date rewritten on an entry that arrived", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const redated = repacked(output, EXTENSIBLE_PATH, (text) =>
      text.replace("2024-04-08T17:32:00Z", "2020-01-01T00:00:00Z")
    );
    expect(onlyCommentsChangedBy(bytes, redated, "u_grace")).toEqual(
      rejected(EXTENSIBLE_PATH)
    );
  });

  it("refuses a date given to a comment that arrived without one", () => {
    const bytes = wordDocx();
    const redated = repacked(bytes, EXTENSIBLE_PATH, (text) =>
      text.replace(
        "</w16cex:commentsExtensible>",
        `${dateEntry("483A933D", "2020-01-01T00:00:00Z")}</w16cex:commentsExtensible>`
      )
    );
    expect(onlyCommentsChangedBy(bytes, redated, "u_grace")).toEqual(
      rejected(EXTENSIBLE_PATH)
    );
  });

  it("refuses a durable id re-pointed on an entry that arrived", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const repointed = repacked(output, IDS_PATH, (text) =>
      text.replace('w16cid:durableId="483A933D"', 'w16cid:durableId="483A933E"')
    );
    expect(onlyCommentsChangedBy(bytes, repointed, "u_grace")).toEqual(
      rejected(IDS_PATH)
    );
  });

  it("refuses the durable id of a comment still standing taken out", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const undated = repacked(output, IDS_PATH, (text) =>
      text.replace(idEntry("1A2B3C4D", "12EC8C19"), "")
    );
    expect(onlyCommentsChangedBy(bytes, undated, "u_grace")).toEqual(
      rejected(IDS_PATH)
    );
  });

  it("refuses a new comment dated further from its wall clock than any zone lies from UTC", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const moved = (dateUtc: string) =>
      repacked(output, EXTENSIBLE_PATH, (text) =>
        text.replace(
          'w16cex:dateUtc="2026-09-30T06:00:00Z"',
          `w16cex:dateUtc="${dateUtc}"`
        )
      );

    expect(
      onlyCommentsChangedBy(bytes, moved("2026-09-29T06:00:00Z"), "u_grace")
    ).toEqual(rejected(EXTENSIBLE_PATH));
    // A clock fourteen hours behind UTC could have written the same wall clock
    expect(
      onlyCommentsChangedBy(bytes, moved("2026-10-01T05:00:00Z"), "u_grace")
    ).toEqual({ ok: true });
  });

  it("refuses a new durable id that one arrived under already", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const { durableId } = writtenIds(output, "4");
    const reused = repacked(
      repacked(output, IDS_PATH, (text) =>
        text.replace(
          `w16cid:durableId="${durableId}"`,
          'w16cid:durableId="483A933D"'
        )
      ),
      EXTENSIBLE_PATH,
      (text) =>
        text.replace(
          `w16cex:durableId="${durableId}"`,
          'w16cex:durableId="483A933D"'
        )
    );
    expect(onlyCommentsChangedBy(bytes, reused, "u_grace")).toEqual(
      rejected(IDS_PATH)
    );
  });

  it("refuses markup the writer does not put on an entry", () => {
    const bytes = wordDocx();
    const output = commentedByGrace(bytes);
    const { durableId } = writtenIds(output, "4");
    const forged = repacked(output, EXTENSIBLE_PATH, (text) =>
      text.replace(
        `w16cex:durableId="${durableId}"`,
        `w16cex:durableId="${durableId}" w16cex:intelligentPlaceholder="1"`
      )
    );
    expect(onlyCommentsChangedBy(bytes, forged, "u_grace")).toEqual(
      rejected(EXTENSIBLE_PATH)
    );
  });
});
