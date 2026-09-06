// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { decode, makeDocx } from "../../__testing__/docx";
import { rangeOfText } from "../../__testing__/editing";
import {
  addComment,
  addCommentReply,
  setCommentResolved,
  updateComment,
} from "../../editor/commands/commentCommands";
import { createEditorState } from "../../editor/createEditor";
import { parseXml } from "../../ooxml/xml";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import { W14_NS, W15_NS } from "./constants";
import { commentPartsKept, entryAllowed, wellFormedEntry } from "./verifying";

const encoder = new TextEncoder();
const COMMENTS_PART = "word/comments.xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

function contentTypes(overrides: string): string {
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    overrides +
    "</Types>"
  );
}

/** A plain document a comment and its author can be added to */
function plainDocx(body = `<w:p>${run("Alpha beta")}</w:p>`): Uint8Array {
  const parts = unzipSync(makeDocx(body));
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

/** The file before and after this author commented on "beta" through the editor */
function commentedBy(
  authorId: string,
  author = "Someone"
): { bytes: Uint8Array; commented: Uint8Array } {
  const bytes = plainDocx();
  const { doc, session } = importDocx(bytes);
  const state = applied(
    selecting(createEditorState(doc), "beta"),
    addComment({ text: "note", author, authorId })
  );
  return { bytes, commented: exportDocx(state.doc, session) };
}

/** The file after this author commented and then settled the thread, which writes the extended part */
function settledByMe(): Uint8Array {
  const bytes = plainDocx();
  const { doc, session } = importDocx(bytes);
  const added = applied(
    selecting(createEditorState(doc), "beta"),
    addComment({ text: "note", author: "Someone", authorId: "me" })
  );
  const settled = applied(added, setCommentResolved("0", true));
  return exportDocx(settled.doc, session);
}

function partText(bytes: Uint8Array, path: string): string {
  return decode(unzipSync(bytes)[path]);
}

function repacked(bytes: Uint8Array, path: string, text: string): Uint8Array {
  const parts = unzipSync(bytes);
  parts[path] = encoder.encode(text);
  return zipSync(parts);
}

const verdict = (
  before: Uint8Array,
  after: Uint8Array,
  authorId: string,
  editableComments: "own" | "all" = "own"
) =>
  commentPartsKept(
    importDocx(before),
    importDocx(after),
    authorId,
    editableComments
  );

const allowed = { ok: true };
const refused = { ok: false, reason: "part-changed", part: COMMENTS_PART };

const elementOf = (xml: string): Element => parseXml(xml).documentElement;

const comment = (attrs: string, body: string): Element =>
  elementOf(
    `<w:comment xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" ${attrs}>${body}</w:comment>`
  );

const BODY = '<w:p><w:r><w:t xml:space="preserve">note</w:t></w:r></w:p>';

const NO_PEOPLE = {
  partPath: null,
  xml: null,
  hadBom: false,
  byAuthor: new Map(),
};

describe("the shape this editor writes", () => {
  it("takes a body written the way the writer writes it", () => {
    const { commented } = commentedBy("me");
    const entry = elementOf(partText(commented, COMMENTS_PART)).children[0];

    expect(wellFormedEntry(entry)).toBe(true);
  });

  it("takes a body carrying line breaks and an empty one", () => {
    for (const body of [
      '<w:p><w:r><w:t xml:space="preserve">a</w:t><w:br/><w:t xml:space="preserve">b</w:t></w:r></w:p>',
      '<w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>',
      `<w:p w14:paraId="12345678"><w:r><w:t xml:space="preserve">a</w:t></w:r></w:p>`,
    ]) {
      expect(wellFormedEntry(comment('w:id="0"', body))).toBe(true);
    }
  });

  it("refuses a body holding anything the writer does not put there", () => {
    for (const body of [
      // a field, a second run, a run property, a tab, two paragraphs, a wrapper
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>',
      `<w:p>${run("a")}${run("b")}</w:p>`,
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p>",
      "<w:p><w:r><w:tab/></w:r></w:p>",
      `<w:p>${run("a")}</w:p><w:p>${run("b")}</w:p>`,
      `<w:p><w:hyperlink>${run("a")}</w:hyperlink></w:p>`,
    ]) {
      expect(wellFormedEntry(comment('w:id="0"', body))).toBe(false);
    }
  });

  it("refuses an attribute this editor does not write on a comment", () => {
    for (const attrs of [
      'w:id="0" w:done="1"',
      'w:id="0" w14:paraId="12345678"',
    ]) {
      expect(wellFormedEntry(comment(attrs, BODY))).toBe(false);
    }
  });

  it("refuses bytes riding between the pieces of a body", () => {
    for (const body of [
      '<w:p><!-- payload --><w:r><w:t xml:space="preserve">a</w:t></w:r></w:p>',
      '<w:p>payload<w:r><w:t xml:space="preserve">a</w:t></w:r></w:p>',
      '<w:p><w:r>payload<w:t xml:space="preserve">a</w:t></w:r></w:p>',
      '<w:p><w:r><w:t xml:space="preserve">a<!-- payload --></w:t></w:r></w:p>',
    ]) {
      expect(wellFormedEntry(comment('w:id="0"', body))).toBe(false);
    }
  });

  it("refuses a w:t that does not keep its space, which the writer always says", () => {
    expect(
      wellFormedEntry(comment('w:id="0"', "<w:p><w:r><w:t>a</w:t></w:r></w:p>"))
    ).toBe(false);
  });

  const extension = (attrs: string): Element =>
    elementOf(`<w15:commentEx xmlns:w15="${W15_NS}" ${attrs}/>`);

  const person = (inner: string, attrs = 'w15:author="Someone"'): Element =>
    elementOf(
      `<w15:person xmlns:w15="${W15_NS}" ${attrs}>${inner}</w15:person>`
    );

  const PRESENCE =
    '<w15:presenceInfo w15:providerId="portone-docx-editor" w15:userId="me"/>';

  it("takes the thread state and the identity this editor writes", () => {
    expect(
      wellFormedEntry(extension('w15:paraId="12345678" w15:done="1"'))
    ).toBe(true);
    expect(wellFormedEntry(person(PRESENCE))).toBe(true);
  });

  it("refuses thread state this editor does not write", () => {
    for (const entry of [
      extension('w15:paraId="12345678" w15:resolved="1"'),
      elementOf(
        `<w15:commentEx xmlns:w15="${W15_NS}" w15:paraId="12345678"><w15:extra/></w15:commentEx>`
      ),
    ]) {
      expect(wellFormedEntry(entry)).toBe(false);
    }
  });

  it("refuses an identity this editor did not record", () => {
    for (const entry of [
      person(""),
      person(PRESENCE + PRESENCE),
      person('<w15:presenceInfo w15:providerId="AD" w15:userId="me"/>'),
      person('<w15:presenceInfo w15:providerId="portone-docx-editor"/>'),
      person(PRESENCE, 'w15:author="Someone" w15:extra="1"'),
    ]) {
      expect(wellFormedEntry(entry)).toBe(false);
    }
  });
});

describe("who may have written an entry", () => {
  const people = {
    partPath: "word/people.xml",
    xml: null,
    hadBom: false,
    byAuthor: new Map([
      ["Mine", "me"],
      ["Theirs", "other"],
    ]),
  };

  it("takes a comment that appeared under this author's name", () => {
    const entry = comment('w:id="0" w:author="Mine"', BODY);

    expect(entryAllowed(entry, null, "me", "own", people)).toBe(true);
    expect(entryAllowed(entry, null, "other", "own", people)).toBe(false);
  });

  it("allows a moderator's rewrite of another author's body under all", () => {
    const original = comment('w:id="0" w:author="Theirs"', BODY);
    const rewritten = comment(
      'w:id="0" w:author="Theirs"',
      '<w:p><w:r><w:t xml:space="preserve">rewritten</w:t></w:r></w:p>'
    );

    expect(entryAllowed(rewritten, original, "me", "own", people)).toBe(false);
    expect(entryAllowed(rewritten, original, "me", "all", people)).toBe(true);
  });

  it("refuses a rewritten w:author or w:date on an entry that arrived, under all too", () => {
    const original = comment(
      'w:id="0" w:author="Mine" w:date="2020-01-01T00:00:00Z"',
      BODY
    );
    for (const attrs of [
      'w:id="0" w:author="Theirs" w:date="2020-01-01T00:00:00Z"',
      'w:id="0" w:author="Mine" w:date="2021-01-01T00:00:00Z"',
      'w:id="1" w:author="Mine" w:date="2020-01-01T00:00:00Z"',
    ]) {
      const rewritten = comment(attrs, BODY);
      expect(entryAllowed(rewritten, original, "me", "own", people)).toBe(
        false
      );
      expect(entryAllowed(rewritten, original, "me", "all", people)).toBe(
        false
      );
    }
  });

  it("takes a thread key that appears, and refuses one re-pointed", () => {
    const threaded = (paraId: string | null) =>
      comment(
        'w:id="0" w:author="Mine"',
        paraId === null
          ? BODY
          : `<w:p w14:paraId="${paraId}"><w:r><w:t xml:space="preserve">note</w:t></w:r></w:p>`
      );

    // Settling a comment for the first time is where the key gets written
    expect(
      entryAllowed(threaded("11111111"), threaded(null), "me", "own", people)
    ).toBe(true);
    expect(
      entryAllowed(
        threaded("22222222"),
        threaded("11111111"),
        "me",
        "all",
        people
      )
    ).toBe(false);
  });

  it("refuses a person recorded for another identity", () => {
    const person = (userId: string) =>
      elementOf(
        `<w15:person xmlns:w15="${W15_NS}" w15:author="Mine">` +
          `<w15:presenceInfo w15:providerId="portone-docx-editor" w15:userId="${userId}"/>` +
          "</w15:person>"
      );

    expect(entryAllowed(person("me"), null, "me", "own", NO_PEOPLE)).toBe(true);
    expect(entryAllowed(person("other"), null, "me", "own", NO_PEOPLE)).toBe(
      false
    );
    // An identity already recorded is nobody's to rewrite, its own subject included
    expect(
      entryAllowed(person("me"), person("other"), "me", "all", NO_PEOPLE)
    ).toBe(false);
  });

  it("refuses a commentEx re-threaded onto somebody else's paragraph", () => {
    const extension = (attrs: string) =>
      elementOf(`<w15:commentEx xmlns:w15="${W15_NS}" ${attrs}/>`);
    const original = extension('w15:paraId="11111111" w15:done="0"');

    expect(
      entryAllowed(
        extension('w15:paraId="11111111" w15:done="1"'),
        original,
        "me",
        "own",
        NO_PEOPLE
      )
    ).toBe(true);
    expect(
      entryAllowed(
        extension('w15:paraId="11111111" w15:paraIdParent="22222222"'),
        original,
        "me",
        "own",
        NO_PEOPLE
      )
    ).toBe(false);
  });
});

describe("over the comment parts of a submitted file", () => {
  it("holds for the file the editor wrote", () => {
    const { bytes, commented } = commentedBy("me");

    expect(verdict(bytes, commented, "me")).toEqual(allowed);
  });

  it("refuses an INCLUDEPICTURE field forged into a comment", () => {
    const { bytes, commented } = commentedBy("me");
    const forged = partText(commented, COMMENTS_PART).replace(
      "</w:p></w:comment>",
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "http://evil.example/x.png" </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:comment>'
    );

    expect(
      verdict(bytes, repacked(commented, COMMENTS_PART, forged), "me")
    ).toEqual(refused);
  });

  it("refuses mc:AlternateContent inside one's own comment", () => {
    const { bytes, commented } = commentedBy("me");
    const wrapped = partText(commented, COMMENTS_PART).replace(
      "</w:p></w:comment>",
      `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Fallback>${run("x")}</mc:Fallback></mc:AlternateContent></w:p></w:comment>`
    );

    expect(
      verdict(bytes, repacked(commented, COMMENTS_PART, wrapped), "me")
    ).toEqual(refused);
  });

  it("refuses an orphan w:comment that appeared under a third author", () => {
    const { bytes, commented } = commentedBy("me");
    const orphaned = partText(commented, COMMENTS_PART).replace(
      "</w:comments>",
      '<w:comment w:id="999" w:author="Someone Else" w:date="2020-01-01T00:00:00Z">' +
        `<w:p>${run("ghost")}</w:p></w:comment></w:comments>`
    );

    expect(
      verdict(bytes, repacked(commented, COMMENTS_PART, orphaned), "me")
    ).toEqual(refused);
  });

  it("keeps an orphan that arrived with the file, and refuses one that changed", () => {
    const { commented } = commentedBy("me");
    const orphan =
      '<w:comment w:id="999" w:author="Someone Else" w:date="2020-01-01T00:00:00Z">' +
      `<w:p>${run("ghost")}</w:p></w:comment>`;
    const withOrphan = repacked(
      commented,
      COMMENTS_PART,
      partText(commented, COMMENTS_PART).replace(
        "</w:comments>",
        `${orphan}</w:comments>`
      )
    );
    const rewritten = repacked(
      withOrphan,
      COMMENTS_PART,
      partText(withOrphan, COMMENTS_PART).replace("ghost", "haunted")
    );
    const dropped = repacked(
      withOrphan,
      COMMENTS_PART,
      partText(withOrphan, COMMENTS_PART).replace(orphan, "")
    );

    expect(verdict(withOrphan, withOrphan, "me")).toEqual(allowed);
    expect(verdict(withOrphan, rewritten, "me")).toEqual(refused);
    expect(verdict(withOrphan, dropped, "me")).toEqual(refused);
  });

  /**
   * The extended part stands for comments by their thread key, so an entry keyed to no comment is
   * one no edit through the editor could have written, and one it stood behind cannot go away.
   */
  it("holds the extended comments part to the same rule", () => {
    const settled = settledByMe();
    const extendedPath = Object.keys(unzipSync(settled)).find((path) =>
      path.endsWith("commentsExtended.xml")
    );
    if (extendedPath === undefined) throw new Error("no extended part");
    const refusedThere = {
      ok: false,
      reason: "part-changed",
      part: extendedPath,
    };
    const ghost = '<w15:commentEx w15:paraId="DEADBEEF" w15:done="1"/>';
    const withGhost = repacked(
      settled,
      extendedPath,
      partText(settled, extendedPath).replace(
        "</w15:commentsEx>",
        `${ghost}</w15:commentsEx>`
      )
    );
    const dropped = repacked(
      withGhost,
      extendedPath,
      partText(withGhost, extendedPath).replace(ghost, "")
    );

    expect(verdict(withGhost, withGhost, "me")).toEqual(allowed);
    // Thread state keyed to no comment is state no edit through the editor could have written
    expect(verdict(settled, withGhost, "me")).toEqual(refusedThere);
    expect(verdict(withGhost, dropped, "me")).toEqual(refusedThere);
  });

  /** The same rule over the other two parts: thread state for no comment, an identity for no name */
  it("keeps an orphan of the other two parts as well", () => {
    const { commented } = commentedBy("me");
    const peoplePath = Object.keys(unzipSync(commented)).find((path) =>
      path.endsWith("people.xml")
    );
    if (peoplePath === undefined) throw new Error("no people part");
    const ghost =
      '<w15:person w15:author="Ghost">' +
      '<w15:presenceInfo w15:providerId="portone-docx-editor" w15:userId="ghost"/>' +
      "</w15:person>";
    const withGhost = repacked(
      commented,
      peoplePath,
      partText(commented, peoplePath).replace(
        "</w15:people>",
        `${ghost}</w15:people>`
      )
    );
    const dropped = repacked(
      withGhost,
      peoplePath,
      partText(withGhost, peoplePath).replace(ghost, "")
    );

    expect(verdict(withGhost, withGhost, "me")).toEqual(allowed);
    expect(verdict(withGhost, dropped, "me")).toEqual({
      ok: false,
      reason: "part-changed",
      part: peoplePath,
    });
  });

  it("refuses an identity recorded for a name nobody writes under", () => {
    const { commented } = commentedBy("me");
    const peoplePath = Object.keys(unzipSync(commented)).find((path) =>
      path.endsWith("people.xml")
    );
    if (peoplePath === undefined) throw new Error("no people part");
    const added = repacked(
      commented,
      peoplePath,
      partText(commented, peoplePath).replace(
        "</w15:people>",
        '<w15:person w15:author="Nobody">' +
          '<w15:presenceInfo w15:providerId="portone-docx-editor" w15:userId="me"/>' +
          "</w15:person></w15:people>"
      )
    );

    expect(verdict(commented, added, "me")).toEqual({
      ok: false,
      reason: "part-changed",
      part: peoplePath,
    });
  });

  it("refuses a second entry smuggled in under an id the part already holds", () => {
    const { bytes, commented } = commentedBy("me");
    const text = partText(commented, COMMENTS_PART);
    const twice = text.replace(
      "</w:comments>",
      `<w:comment w:id="0" w:author="Someone"><w:p>${run("payload")}</w:p></w:comment></w:comments>`
    );

    expect(
      verdict(bytes, repacked(commented, COMMENTS_PART, twice), "me")
    ).toEqual(refused);
  });

  /**
   * The judgement has to take everything the editor itself writes, or a file a reader made in good
   * faith is turned down. Each command is run over a comment of this author's, over one written by
   * somebody else, and over one carrying no recorded identity, which is the shape a document that
   * arrived from Word has and the one `editableComments: "own"` calls everyone's.
   */
  describe("what the editing commands write", () => {
    const ME = { id: "me", name: "Someone" };
    const REPLY = { text: "reply", author: "Someone", authorId: "me" };

    const edits: readonly (readonly [string, Command])[] = [
      ["rewritten body", updateComment("0", "rewritten")],
      ["settled thread", setCommentResolved("0", true)],
      ["added reply", addCommentReply("0", REPLY)],
    ];

    /** The file with a comment whose author this editor recorded no identity for */
    function commentedAnonymously(): Uint8Array {
      const { commented } = commentedBy("them", "Somebody Else");
      const people = Object.keys(unzipSync(commented)).find((path) =>
        path.endsWith("people.xml")
      );
      if (people === undefined) throw new Error("no people part");
      const parts = unzipSync(commented);
      delete parts[people];
      return zipSync(parts);
    }

    function holds(before: Uint8Array, editableComments: "own" | "all") {
      const { doc, session } = importDocx(before);
      const opened = createEditorState(doc, { author: ME, editableComments });
      for (const [what, command] of edits) {
        const state = applied(opened, command);
        const submitted = exportDocx(state.doc, session);
        expect([
          what,
          verdict(before, submitted, "me", editableComments),
        ]).toEqual([what, allowed]);
      }
    }

    it("holds for a comment of this author's", () => {
      holds(commentedBy("me").commented, "own");
    });

    it("holds for a comment carrying no recorded identity", () => {
      holds(commentedAnonymously(), "own");
    });

    it("holds for another author's comment under `all`", () => {
      holds(commentedBy("them", "Somebody Else").commented, "all");
    });

    it("does not take another author's body rewritten under `own`", () => {
      const before = commentedBy("them", "Somebody Else").commented;
      const { doc, session } = importDocx(before);
      const state = applied(
        createEditorState(doc, { editableComments: "all" }),
        updateComment("0", "rewritten")
      );

      expect(verdict(before, exportDocx(state.doc, session), "me")).toEqual(
        refused
      );
    });

    it("takes another author's thread settled under `own`, which is everyone's", () => {
      const before = commentedBy("them", "Somebody Else").commented;
      const { doc, session } = importDocx(before);
      const state = applied(
        createEditorState(doc, { author: ME }),
        setCommentResolved("0", true)
      );

      expect(verdict(before, exportDocx(state.doc, session), "me")).toEqual(
        allowed
      );
    });
  });
});
