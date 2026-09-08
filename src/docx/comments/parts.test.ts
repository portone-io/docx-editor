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
import { onlyCommentsChangedBy } from "../commentOnlyChange";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import { partsKept } from "../protectionPolicy";
import { W14_NS, W15_NS } from "./constants";
import { entryAllowed, wellFormedEntry } from "./parts";
import { commentsPolicy } from "./policy";

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
  partsKept(commentsPolicy, importDocx(before), importDocx(after), authorId, {
    editableComments,
  });

const allowed = { ok: true };
const refused = {
  ok: false,
  reason: "comment-markup-rejected",
  part: COMMENTS_PART,
};

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

  /**
   * A declaration decides what every name under it means, so a rebound prefix is a body that says
   * one thing to this reader and another to Word.
   */
  it("refuses a namespace declaration the writer does not write", () => {
    for (const entry of [
      elementOf(
        `<w:comment xmlns:w="${W_NS}" xmlns:evil="urn:x" w:id="0">${BODY}</w:comment>`
      ),
      elementOf(
        `<w:comment xmlns:w="${W_NS}" w:id="0">` +
          '<w:p xmlns:w14="urn:not-word" w14:paraId="12345678">' +
          '<w:r><w:t xml:space="preserve">note</w:t></w:r></w:p></w:comment>'
      ),
      elementOf(
        `<w15:person xmlns:w15="${W15_NS}" w15:author="Someone">` +
          '<w15:presenceInfo xmlns:w15="urn:not-word"' +
          ' w15:providerId="portone-docx-editor" w15:userId="me"/></w15:person>'
      ),
    ]) {
      expect(wellFormedEntry(entry)).toBe(false);
    }
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
  /** The names a document already comments under while recording nobody for them */
  const NAMED_BY_NOBODY: ReadonlySet<string> = new Set(["Nobody's"]);

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

    expect(
      entryAllowed(entry, null, "me", "own", people, NAMED_BY_NOBODY)
    ).toBe(true);
    expect(
      entryAllowed(entry, null, "other", "own", people, NAMED_BY_NOBODY)
    ).toBe(false);
  });

  it("takes a comment that appeared naming nobody, under a name the original already leaves unattributed", () => {
    const named = comment('w:id="0" w:author="Nobody\'s"', BODY);
    const stranger = comment('w:id="0" w:author="Passing"', BODY);

    // No person is recorded for either name, so both resolve to no identity here. What tells them
    // apart is that the original already showed one of them and never showed the other
    expect(
      entryAllowed(named, null, "me", "own", NO_PEOPLE, NAMED_BY_NOBODY)
    ).toBe(true);
    expect(
      entryAllowed(stranger, null, "me", "own", NO_PEOPLE, NAMED_BY_NOBODY)
    ).toBe(false);
    expect(entryAllowed(named, null, "me", "own", NO_PEOPLE, new Set())).toBe(
      false
    );
  });

  it("allows a moderator's rewrite of another author's body under all", () => {
    const original = comment('w:id="0" w:author="Theirs"', BODY);
    const rewritten = comment(
      'w:id="0" w:author="Theirs"',
      '<w:p><w:r><w:t xml:space="preserve">rewritten</w:t></w:r></w:p>'
    );

    expect(
      entryAllowed(rewritten, original, "me", "own", people, NAMED_BY_NOBODY)
    ).toBe(false);
    expect(
      entryAllowed(rewritten, original, "me", "all", people, NAMED_BY_NOBODY)
    ).toBe(true);
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
      expect(
        entryAllowed(rewritten, original, "me", "own", people, NAMED_BY_NOBODY)
      ).toBe(false);
      expect(
        entryAllowed(rewritten, original, "me", "all", people, NAMED_BY_NOBODY)
      ).toBe(false);
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
      entryAllowed(
        threaded("11111111"),
        threaded(null),
        "me",
        "own",
        people,
        NAMED_BY_NOBODY
      )
    ).toBe(true);
    expect(
      entryAllowed(
        threaded("22222222"),
        threaded("11111111"),
        "me",
        "all",
        people,
        NAMED_BY_NOBODY
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

    expect(
      entryAllowed(person("me"), null, "me", "own", NO_PEOPLE, NAMED_BY_NOBODY)
    ).toBe(true);
    expect(
      entryAllowed(
        person("me"),
        null,
        "me",
        "own",
        NO_PEOPLE,
        new Set(["Mine"])
      )
    ).toBe(false);
    expect(
      entryAllowed(
        person("other"),
        null,
        "me",
        "own",
        NO_PEOPLE,
        NAMED_BY_NOBODY
      )
    ).toBe(false);
    // An identity already recorded is nobody's to rewrite, its own subject included
    expect(
      entryAllowed(
        person("me"),
        person("other"),
        "me",
        "all",
        NO_PEOPLE,
        NAMED_BY_NOBODY
      )
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
        NO_PEOPLE,
        NAMED_BY_NOBODY
      )
    ).toBe(true);
    expect(
      entryAllowed(
        extension('w15:paraId="11111111" w15:paraIdParent="22222222"'),
        original,
        "me",
        "own",
        NO_PEOPLE,
        NAMED_BY_NOBODY
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
      reason: "comment-markup-rejected",
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
      reason: "comment-markup-rejected",
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
      reason: "comment-markup-rejected",
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
   * The entries are judged one by one, so the room around them is where bytes go that no entry
   * judgement ever reads. Whitespace is the one thing that says nothing, and a part that arrived
   * laid out over several lines comes back that way.
   */
  it("refuses bytes riding between the entries of a part, and takes the layout of one", () => {
    const { bytes, commented } = commentedBy("me");
    const text = partText(commented, COMMENTS_PART);
    for (const between of [
      "<!-- payload -->",
      "<?review payload?>",
      "\u00a0",
      "payload",
      "<![CDATA[payload]]>",
    ]) {
      const ridden = text.replace("</w:comments>", `${between}</w:comments>`);
      expect([
        between,
        verdict(bytes, repacked(commented, COMMENTS_PART, ridden), "me"),
      ]).toEqual([between, refused]);
    }

    const laidOut = repacked(
      commented,
      COMMENTS_PART,
      text.replace("</w:comments>", "\n  </w:comments>")
    );
    expect(verdict(laidOut, laidOut, "me")).toEqual(allowed);
  });

  it("refuses a part root the file did not arrive with", () => {
    const settled = settledByMe();
    const text = partText(settled, COMMENTS_PART);
    const roots: readonly (readonly [string, string])[] = [
      [
        "a rebound prefix",
        text.replace(`xmlns:w14="${W14_NS}"`, 'xmlns:w14="urn:x"'),
      ],
      [
        "a declaration the writer never adds",
        text.replace(
          `xmlns:w="${W_NS}"`,
          `xmlns:w="${W_NS}" xmlns:evil="urn:x"`
        ),
      ],
      [
        "another name passed over",
        text.replace('mc:Ignorable="w14"', 'mc:Ignorable="w14 evil"'),
      ],
    ];
    for (const [what, root] of roots) {
      expect([
        what,
        verdict(settled, repacked(settled, COMMENTS_PART, root), "me"),
      ]).toEqual([what, refused]);
    }

    // and a declaration the file arrived with cannot go away either
    const declared = repacked(
      settled,
      COMMENTS_PART,
      text.replace(`xmlns:w="${W_NS}"`, `xmlns:w="${W_NS}" xmlns:extra="urn:x"`)
    );
    expect(verdict(declared, settled, "me")).toEqual(refused);
  });

  /** Settling a thread is where the writer adds the compatibility markup to a root that arrived */
  it("takes the compatibility markup the writer adds to a part it settles a thread in", () => {
    const { commented } = commentedBy("me");
    const { doc, session } = importDocx(commented);
    const settled = applied(
      createEditorState(doc),
      setCommentResolved("0", true)
    );
    const after = exportDocx(settled.doc, session);

    expect(partText(commented, COMMENTS_PART)).not.toContain("mc:Ignorable");
    expect(partText(after, COMMENTS_PART)).toContain('mc:Ignorable="w14"');
    expect(verdict(commented, after, "me")).toEqual(allowed);
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

describe("the content surrounding comment entries", () => {
  const paths = [COMMENTS_PART, "word/commentsExtended.xml", "word/people.xml"];
  const annotation = "<!-- producer annotation -->";
  const refusedAt = (part: string) => ({
    ok: false,
    reason: "comment-markup-rejected",
    part,
  });

  function annotated(bytes: Uint8Array, path: string): Uint8Array {
    const xml = partText(bytes, path);
    const close = xml.lastIndexOf("</");
    return repacked(
      bytes,
      path,
      xml.slice(0, close) + annotation + xml.slice(close)
    );
  }

  it.each(paths)(
    "accepts a preserved annotation in %s and refuses its alteration",
    (path) => {
      const before = annotated(settledByMe(), path);
      const { doc, session } = importDocx(before);
      const after = exportDocx(doc, session);
      expect(partText(after, path)).toBe(partText(before, path));
      expect(onlyCommentsChangedBy(before, after, "me")).toEqual(allowed);

      for (const replacement of [
        "<!-- changed annotation -->",
        "",
        annotation + annotation,
      ]) {
        const changed = repacked(
          after,
          path,
          partText(after, path).replace(annotation, replacement)
        );
        expect(onlyCommentsChangedBy(before, changed, "me")).toEqual(
          refusedAt(path)
        );
      }
    }
  );

  it("accepts a reply that appends a person beside an existing annotation", () => {
    const path = "word/people.xml";
    const before = annotated(settledByMe(), path);
    const { doc, session } = importDocx(before);
    const state = applied(
      createEditorState(doc, { author: { id: "other", name: "Another" } }),
      addCommentReply("0", {
        text: "Reply",
        author: "Another",
        authorId: "other",
      })
    );
    const after = exportDocx(state.doc, session);
    expect(partText(after, path)).toContain(annotation);
    expect(partText(after, path)).toContain('w15:userId="other"');
    expect(onlyCommentsChangedBy(before, after, "other")).toEqual(allowed);

    const dropped = repacked(
      after,
      path,
      partText(after, path).replace(annotation, "")
    );
    expect(onlyCommentsChangedBy(before, dropped, "other")).toEqual(
      refusedAt(path)
    );
  });

  it("accepts the inter-entry annotations lost when a reply rebuilds both comment parts", () => {
    const commentPaths = paths.slice(0, 2);
    const before = commentPaths.reduce(annotated, settledByMe());
    const { doc, session } = importDocx(before);
    const state = applied(
      createEditorState(doc, { author: { id: "me", name: "Someone" } }),
      addCommentReply("0", { text: "Reply", author: "Someone", authorId: "me" })
    );
    const after = exportDocx(state.doc, session);
    for (const path of commentPaths)
      expect(partText(after, path)).not.toContain(annotation);
    expect(onlyCommentsChangedBy(before, after, "me")).toEqual(allowed);
  });

  it.each(paths)("compares annotations outside the root of %s", (path) => {
    const original = settledByMe();
    for (const outside of [annotation, "<?review producer-annotation?>"]) {
      const annotatedXml = partText(original, path) + outside;
      const before = repacked(original, path, annotatedXml);
      const { doc, session } = importDocx(before);
      expect(
        onlyCommentsChangedBy(before, exportDocx(doc, session), "me")
      ).toEqual(allowed);
      expect(onlyCommentsChangedBy(original, before, "me")).toEqual(
        refusedAt(path)
      );
      expect(onlyCommentsChangedBy(before, original, "me")).toEqual(
        refusedAt(path)
      );
      const changed = repacked(
        before,
        path,
        annotatedXml.replace("annotation", "changed")
      );
      expect(onlyCommentsChangedBy(before, changed, "me")).toEqual(
        refusedAt(path)
      );
    }
  });
});
