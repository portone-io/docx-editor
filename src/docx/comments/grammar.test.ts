// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml, W_NS } from "../../ooxml/xml";
import { NO_EXPORT_REFS } from "../exportRefs";
import { NO_FORMATTING } from "../formatting";
import { NO_IMPORT_SOURCES } from "../importParagraph";
import { serializeStory } from "../serializeStory";
import { readStory, storyFromText, storyText, withThreadKeyOn } from "../story";
import { W14_NS, W15_NS } from "./constants";
import {
  arrivedEntries,
  recordedIdentity,
  renderCommentExtension,
  renderPerson,
  wellFormedCommentExtension,
  wellFormedPerson,
  withThreadKey,
} from "./grammar";
import { wellFormedEntry } from "./parts";

const CONTAINER = { open: '<w:comment w:id="0">', close: "</w:comment>" };

/** A body this editor writes for that text, under the thread key when there is one */
function written(text: string, key: string | null): string {
  const story = storyFromText(text);
  return serializeStory(
    key === null ? story : withThreadKeyOn(story, key),
    null,
    CONTAINER,
    NO_EXPORT_REFS
  );
}

/** The same body read back as the story it says, the way opening the file reads it */
function readBack(commentXml: string): string {
  const el = parseXml(
    commentXml.replace("<w:comment ", `<w:comment xmlns:w="${W_NS}" `)
  ).documentElement;
  return storyText(
    readStory(
      { kind: "comment", id: "0", partPath: "word/comments.xml" },
      { el, xml: commentXml },
      {
        session: { sessionId: "d0" },
        sources: NO_IMPORT_SOURCES,
        formatting: NO_FORMATTING,
      }
    ).doc
  );
}

/**
 * The writer and the reader are two halves of one grammar, and a file is judged by whether the
 * half that reads recognises what the half that writes put out. Holding them against each other is
 * what keeps them from drifting apart.
 */
describe("what the writer writes, the reader reads back", () => {
  it.each([
    ["plain", "note"],
    ["empty", ""],
    ["a line break", "first\nsecond"],
    ["a blank line", "first\n\nsecond"],
    ["a leading break", "\nsecond"],
    ["a trailing break", "first\n"],
    ["kept space", "  padded  "],
    ["markup characters", 'x < y & z > "w"'],
    ["not latin", "댓글 \u{1F600}"],
  ])("reads %s back as it was written", (_what, text) => {
    for (const key of [null, "12345678"]) {
      const body = written(text, key);
      expect(readBack(body)).toBe(text);
      expect(
        wellFormedEntry(
          parseXml(body.replace("<w:comment ", `<w:comment xmlns:w="${W_NS}" `))
            .documentElement,
          null
        )
      ).toBe(true);
    }
  });

  it("writes a body the reader takes and nothing it does not", () => {
    expect(written("note", null)).toBe(
      '<w:comment w:id="0"><w:p><w:r>' +
        '<w:t xml:space="preserve">note</w:t></w:r></w:p></w:comment>'
    );
    expect(written("a\nb", "12345678")).toBe(
      '<w:comment w:id="0">' +
        `<w:p xmlns:w14="${W14_NS}" w14:paraId="12345678">` +
        '<w:r><w:t xml:space="preserve">a</w:t><w:br/>' +
        '<w:t xml:space="preserve">b</w:t></w:r></w:p></w:comment>'
    );
  });
});

describe("giving an entry a thread key", () => {
  const body = '<w:p><w:r><w:t xml:space="preserve">note</w:t></w:r></w:p>';

  /** The entry as the part has it, which is what says what a prefix means where it is written */
  const partWith = (entry: string, declarations = `xmlns:w="${W_NS}"`) =>
    arrivedEntries(`<w:comments ${declarations}>${entry}</w:comments>`).get(
      "0"
    ) ?? null;

  const keyed = (entry: string, declarations?: string) =>
    withThreadKey(entry, "ABCD1234", partWith(entry, declarations));

  it("puts the key on the body and leaves the rest of the entry alone", () => {
    expect(keyed(`<w:comment w:id="0">${body}</w:comment>`)).toBe(
      '<w:comment w:id="0"><w:p w14:paraId="ABCD1234">' +
        '<w:r><w:t xml:space="preserve">note</w:t></w:r></w:p></w:comment>'
    );
  });

  it("leaves an entry that already carries one exactly as it was", () => {
    const already = `<w:comment w:id="0"><w:p w14:paraId="0000AAAA">${body.slice(5)}</w:comment>`;

    expect(keyed(already, `xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"`)).toBe(
      already
    );
  });

  it("leaves one carrying a key under another prefix alone, rather than writing a second", () => {
    const already = `<w:comment w:id="0"><w:p k:paraId="0000AAAA">${body.slice(5)}</w:comment>`;

    expect(keyed(already, `xmlns:w="${W_NS}" xmlns:k="${W14_NS}"`)).toBe(
      already
    );
  });

  it("puts it on the last paragraph, which is the one the key is read off", () => {
    const two = `<w:comment w:id="0">${body}${body}</w:comment>`;
    const written = keyed(two);

    expect(written.indexOf("w14:paraId")).toBeGreaterThan(
      written.indexOf("</w:p>")
    );
    expect(written.match(/w14:paraId/g)).toHaveLength(1);
  });

  it("keeps a body the writer would not have written, rather than putting one back", () => {
    const rich =
      '<w:comment w:id="0"><w:p><w:r><w:rPr><w:b/></w:rPr>' +
      '<w:t xml:space="preserve">bold</w:t></w:r></w:p></w:comment>';

    expect(keyed(rich)).toContain("<w:b/>");
  });

  it.each(['"', "'"])(
    "takes a paragraph under a second prefix the part binds, declared with %s",
    (quote) => {
      const entry =
        '<w:comment w:id="0"><q:p><q:r><q:t>a</q:t></q:r></q:p></w:comment>';
      const declarations = `xmlns:w=${quote}${W_NS}${quote} xmlns:q=${quote}${W_NS}${quote}`;

      expect(keyed(entry, declarations)).toContain(
        '<q:p w14:paraId="ABCD1234">'
      );
    }
  );

  it("leaves a paragraph of another vocabulary alone", () => {
    const entry =
      '<w:comment w:id="0"><w:p><w:r><w:drawing>' +
      '<a:p xmlns:a="http://example.com/drawing"/></w:drawing></w:r></w:p></w:comment>';

    expect(keyed(entry)).toContain('<w:p w14:paraId="ABCD1234">');
  });

  it("takes a paragraph whose prefix is a name wider than an English one", () => {
    const entry =
      '<w:comment w:id="0"><w\u00f6:p><w\u00f6:r><w\u00f6:t>a</w\u00f6:t></w\u00f6:r></w\u00f6:p></w:comment>';

    expect(keyed(entry, `xmlns:w="${W_NS}" xmlns:w\u00f6="${W_NS}"`)).toContain(
      '<w\u00f6:p w14:paraId="ABCD1234">'
    );
  });

  it("writes no second key beside one spelled the way it would spell its own", () => {
    const entry =
      '<w:comment w:id="0"><w:p w14:paraId="0000AAAA"><w:r><w:t>a</w:t></w:r></w:p></w:comment>';

    expect(
      keyed(entry, `xmlns:w="${W_NS}" xmlns:w14="http://example.com/other"`)
    ).toBe(entry);
  });

  it("reads a prefix as the element it sits on binds it, not as the part first bound it", () => {
    const entry =
      '<w:comment w:id="0"><q:p><q:r><q:drawing>' +
      '<q:p xmlns:q="http://example.com/drawing"/></q:drawing></q:r></q:p></w:comment>';
    const declarations = `xmlns:w="${W_NS}" xmlns:q="${W_NS}"`;

    const written = keyed(entry, declarations);
    expect(written).toContain('<q:p w14:paraId="ABCD1234"><q:r>');
    expect(written).toContain('<q:p xmlns:q="http://example.com/drawing"/>');
  });
});

/** What every comment and reply says about itself, whatever thread state it stands in */
const said = {
  id: "0",
  author: "Someone",
  authorId: "me",
  initials: null,
  date: "2020-01-01T00:00:00Z",
  text: "note",
  commentXml: null,
  imported: false,
  extensionXml: null,
};

const alone = (xml: string): Element =>
  parseXml(`<w15:part xmlns:w15="${W15_NS}">${xml}</w15:part>`).documentElement
    .children[0];

describe("the thread state this editor writes", () => {
  it("is recognised by the half that reads it", () => {
    const settled = renderCommentExtension({
      ...said,
      paraId: "ABCD1234",
      resolved: true,
      threadImported: false,
      replies: [],
    });
    const reply = renderCommentExtension({
      ...said,
      paraId: "0000AAAA",
      parentParaId: "ABCD1234",
    });

    expect(wellFormedCommentExtension(alone(settled))).toBe(true);
    expect(wellFormedCommentExtension(alone(reply))).toBe(true);
  });

  it("is not recognised where a key is not four bytes of hexadecimal", () => {
    const forged = '<w15:commentEx w15:paraId="not a key"/>';

    expect(wellFormedCommentExtension(alone(forged))).toBe(false);
  });
});

describe("the identity this editor records", () => {
  it("is recognised by the half that reads it, and says whose it is", () => {
    const person = renderPerson("Someone", "me", "w15:", null);

    expect(wellFormedPerson(alone(person))).toBe(true);
    expect(recordedIdentity(alone(person))).toBe("me");
  });

  it("is not recognised where another provider recorded it", () => {
    const theirs = renderPerson("Someone", "me", "w15:", null).replace(
      "portone-docx-editor",
      "AD"
    );

    expect(wellFormedPerson(alone(theirs))).toBe(false);
  });
});
