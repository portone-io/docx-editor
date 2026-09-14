// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  makeDeclaredDocx,
  makeNotesDocx,
  withBlocks,
} from "../../__testing__/docx";
import { docxSchema } from "../../schema";
import {
  type StoryKey,
  sameStory,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import { setStory, storyFromText, withoutStories } from "../story";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const FOOTNOTES_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const FOOTNOTES_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const ENDNOTES_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

/** A run of text as the reader models one: every run carries a run mark, its properties or none */
const runText = (text: string, rPr: string | null = null) =>
  docxSchema.text(text, [docxSchema.marks.run.create({ rPr })]);

/** A story of one paragraph per line, as the reader reads back what the writer wrote for it */
const storyOf = (...lines: readonly string[]) =>
  docxSchema.nodes.doc.create(
    null,
    lines.map((line) => docxSchema.nodes.paragraph.create(null, runText(line)))
  );

const reference = (id: string) =>
  `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

const SEPARATOR =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>';
const CONTINUATION =
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

const entry = (id: string, text: string, paraId = "") =>
  `<w:footnote w:id="${id}"><w:p${paraId === "" ? "" : ` w14:paraId="${paraId}"`}>` +
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
  `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`;

/** A package whose Footnotes part holds these entries, laid out the way a producer indents them */
function footnotesDocx(entries: readonly string[], body: string): Uint8Array {
  const parts = unzipSync(makeNotesDocx(body));
  parts["word/footnotes.xml"] = new TextEncoder().encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<w:footnotes xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">\n  ` +
      entries.join("\n  ") +
      "\n</w:footnotes>"
  );
  return zipSync(parts);
}

const FIRST = entry("1", "First note", "10000001");
const SECOND = entry("2", "Second note", "10000002");
const BODY =
  `<w:p><w:r><w:t>One</w:t></w:r>${reference("1")}</w:p>` +
  `<w:p><w:r><w:t>Two</w:t></w:r>${reference("2")}</w:p>`;

function withStories(
  doc: PMNode,
  stories: readonly (readonly [StoryKey, PMNode])[]
): PMNode {
  const tr = new Transform(doc);
  for (const [key, story] of stories) setStory(tr, key, story);
  return tr.doc;
}

/** The footnote keys of a set, leaving out the endnote entries the notes fixture also holds */
function footnoteKeys(keys: Iterable<StoryKey>): readonly StoryKey[] {
  return Array.from(keys).filter((key) => key.startsWith("footnote:"));
}

function writtenPart(bytes: Uint8Array, path: string): string {
  const part = unzipSync(bytes)[path];
  if (part === undefined) throw new Error(`the export wrote no ${path}`);
  return decode(part);
}

describe("writing the footnotes part", () => {
  it("writes back no footnotes part when no footnote story changed", () => {
    const bytes = footnotesDocx([SEPARATOR, FIRST, SECOND], BODY);
    const opened = importDocx(bytes);
    const edited = withBlocks(opened.doc, [
      docxSchema.nodes.paragraph.create(null, docxSchema.text("Body only")),
    ]);

    const after = unzipSync(exportDocx(edited, opened.session));
    const before = unzipSync(bytes);
    expect(
      bytesEqual(after["word/footnotes.xml"], before["word/footnotes.xml"])
    ).toBe(true);
  });

  it("rewrites only the edited footnote and keeps every other entry as it arrived", () => {
    const bytes = footnotesDocx([SEPARATOR, CONTINUATION, FIRST, SECOND], BODY);
    const original = decode(unzipSync(bytes)["word/footnotes.xml"]);
    const opened = importDocx(bytes);
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "1"), storyFromText("Rewritten note")],
    ]);

    const written = writtenPart(
      exportDocx(edited, opened.session),
      "word/footnotes.xml"
    );
    const at = original.indexOf(FIRST);
    expect(written.startsWith(original.slice(0, at))).toBe(true);
    expect(written.endsWith(original.slice(at + FIRST.length))).toBe(true);
    expect(written).toContain(
      '<w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">Rewritten note</w:t></w:r></w:p></w:footnote>'
    );
  });

  it("keeps separator and continuation separator entries through an edit", () => {
    const bytes = footnotesDocx([SEPARATOR, CONTINUATION, FIRST], BODY);
    const opened = importDocx(bytes);
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "1"), storyFromText("Rewritten note")],
    ]);

    const again = importDocx(exportDocx(edited, opened.session));
    const written = decode(
      unzipSync(exportDocx(again.doc, again.session))["word/footnotes.xml"]
    );
    expect(written).toContain(`\n  ${SEPARATOR}\n  ${CONTINUATION}\n  `);
    expect(footnoteKeys(again.session.specialNotes)).toEqual([
      storyKey("footnote", "-1"),
      storyKey("footnote", "0"),
    ]);
  });

  it("leaves out the entry of a footnote whose story was removed", () => {
    const bytes = footnotesDocx([SEPARATOR, FIRST, SECOND], BODY);
    const opened = importDocx(bytes);
    const tr = withoutStories(new Transform(opened.doc), [
      storyKey("footnote", "2"),
    ]);
    const edited = withBlocks(tr.doc, [tr.doc.child(0)]);

    const exported = exportDocx(edited, opened.session);
    const written = writtenPart(exported, "word/footnotes.xml");
    expect(written).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<w:footnotes xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">\n  ` +
        `${SEPARATOR}\n  ${FIRST}\n</w:footnotes>`
    );
    expect(writtenPart(exported, "word/document.xml")).not.toContain(
      'w:footnoteReference w:id="2"'
    );
  });

  it("keeps an entry the file arrived with no reference to", () => {
    const unreferenced = entry("7", "Nobody points here");
    const bytes = footnotesDocx([SEPARATOR, FIRST, unreferenced], BODY);
    const opened = importDocx(bytes);
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "1"), storyFromText("Rewritten note")],
    ]);

    expect(
      writtenPart(exportDocx(edited, opened.session), "word/footnotes.xml")
    ).toContain(`\n  ${unreferenced}\n</w:footnotes>`);
  });

  it("appends a new footnote after the part's own entries", () => {
    const bytes = footnotesDocx([SEPARATOR, FIRST, SECOND], BODY);
    const opened = importDocx(bytes);
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "12"), storyFromText("Twelfth")],
      [storyKey("footnote", "3"), storyFromText("Third")],
    ]);

    const written = writtenPart(
      exportDocx(edited, opened.session),
      "word/footnotes.xml"
    );
    expect(written).toContain(
      `${SECOND}<w:footnote w:id="3"><w:p><w:r><w:t xml:space="preserve">Third</w:t></w:r></w:p></w:footnote>` +
        '<w:footnote w:id="12"><w:p><w:r><w:t xml:space="preserve">Twelfth</w:t></w:r></w:p></w:footnote>\n</w:footnotes>'
    );
  });

  it("creates a footnotes part with its relationship, content type, and separator entries", () => {
    const opened = importDocx(
      makeDeclaredDocx(`<w:p><w:r><w:t>Text</w:t></w:r>${reference("1")}</w:p>`)
    );
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "1"), storyFromText("A new note")],
    ]);

    const exported = exportDocx(edited, opened.session);
    expect(writtenPart(exported, "word/_rels/document.xml.rels")).toMatch(
      new RegExp(
        `<Relationship (?=[^>]*Type="${FOOTNOTES_REL}")(?=[^>]*Target="footnotes.xml")`
      )
    );
    expect(writtenPart(exported, "[Content_Types].xml")).toContain(
      `<Override PartName="/word/footnotes.xml" ContentType="${FOOTNOTES_TYPE}"/>`
    );
    const written = writtenPart(exported, "word/footnotes.xml");
    expect(written).toMatch(
      /^<\?xml [^>]*\?><w:footnotes [^>]*><w:footnote w:type="separator" w:id="-1">.*<w:separator\/>.*<\/w:footnote><w:footnote w:type="continuationSeparator" w:id="0">.*<w:continuationSeparator\/>.*<\/w:footnote><w:footnote w:id="1">/
    );
    const again = importDocx(exported);
    expect(
      sameStory(
        storyNodeOf(again.doc, storyKey("footnote", "1")),
        storyOf("A new note")
      )
    ).toBe(true);
    expect(footnoteKeys(again.session.specialNotes)).toEqual([
      storyKey("footnote", "-1"),
      storyKey("footnote", "0"),
    ]);
  });

  it("gives the new separator entries ids no reference already uses", () => {
    const opened = importDocx(
      makeDeclaredDocx(
        `<w:p><w:r><w:t>Text</w:t></w:r>${reference("0")}${reference("-1")}</w:p>`
      )
    );
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "0"), storyFromText("Note zero")],
    ]);

    const written = writtenPart(
      exportDocx(edited, opened.session),
      "word/footnotes.xml"
    );
    expect(written).toContain('<w:footnote w:type="separator" w:id="-2">');
    expect(written).toContain(
      '<w:footnote w:type="continuationSeparator" w:id="-3">'
    );
    const ids = Array.from(written.matchAll(/ w:id="([^"]*)"/g), (m) => m[1]);
    expect(ids).toEqual(["-2", "-3", "0"]);
  });

  it("writes an edited footnote whose pasted paragraph repeats another footnote's paragraph id without the repeat", () => {
    const bytes = footnotesDocx([SEPARATOR, FIRST, SECOND], BODY);
    const opened = importDocx(bytes);
    const second = storyNodeOf(opened.doc, storyKey("footnote", "2"));
    const first = storyNodeOf(opened.doc, storyKey("footnote", "1"));
    if (!second?.firstChild || !first)
      throw new Error("the notes were not read");
    const pasted = second.firstChild.type.create(second.firstChild.attrs, [
      docxSchema.text("Pasted"),
    ]);
    const edited = withStories(opened.doc, [
      [
        storyKey("footnote", "1"),
        first.copy(first.content.append(Fragment.from(pasted))),
      ],
    ]);

    const written = writtenPart(
      exportDocx(edited, opened.session),
      "word/footnotes.xml"
    );
    expect(written).toContain("Pasted");
    expect(written).toContain("Second note");
    expect(written.match(/w14:paraId="10000002"/g)).toHaveLength(1);
  });

  it("keeps a separator entry's paragraph id and releases it from the footnote that was edited", () => {
    const separator =
      '<w:footnote w:type="separator" w:id="-1"><w:p w14:paraId="10000001">' +
      "<w:r><w:separator/></w:r></w:p></w:footnote>";
    const bytes = footnotesDocx(
      [FIRST, separator],
      `<w:p><w:r><w:t>One</w:t></w:r>${reference("1")}</w:p>`
    );
    const opened = importDocx(bytes);
    const first = storyNodeOf(opened.doc, storyKey("footnote", "1"));
    if (first === null) throw new Error("the note was not read");
    const edited = withStories(opened.doc, [
      [
        storyKey("footnote", "1"),
        first.copy(
          first.content.append(
            Fragment.from(
              docxSchema.nodes.paragraph.create(null, runText("Appended"))
            )
          )
        ),
      ],
    ]);

    const written = writtenPart(
      exportDocx(edited, opened.session),
      "word/footnotes.xml"
    );
    expect(written).toContain(separator);
    expect(written.match(/w14:paraId="10000001"/g)).toHaveLength(1);
  });

  it("reads an edited footnote back as the story that was written", () => {
    const bytes = footnotesDocx([SEPARATOR, FIRST, SECOND], BODY);
    const opened = importDocx(bytes);
    const story = docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create(
        { pPr: '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' },
        [runText("Bold", "<w:rPr><w:b/></w:rPr>"), runText(" plain")]
      ),
      docxSchema.nodes.paragraph.create(null, runText("Second line")),
    ]);
    const key = storyKey("footnote", "2");
    const edited = withStories(opened.doc, [[key, story]]);

    const again = importDocx(exportDocx(edited, opened.session));
    expect(sameStory(storyNodeOf(again.doc, key), story)).toBe(true);
  });
});

describe("writing the endnotes part", () => {
  it("rewrites only the edited endnote and keeps every other entry as it arrived", () => {
    const bytes = makeNotesDocx();
    const original = decode(unzipSync(bytes)["word/endnotes.xml"]);
    const opened = importDocx(bytes);
    const edited = withStories(opened.doc, [
      [storyKey("endnote", "3"), storyFromText("Rewritten endnote")],
    ]);

    const exported = exportDocx(edited, opened.session);
    const written = writtenPart(exported, "word/endnotes.xml");
    const continuation = original.slice(
      original.indexOf("<w:endnote "),
      original.indexOf('<w:endnote w:id="3">')
    );
    expect(written).toContain(continuation);
    expect(written).toContain(
      '<w:endnote w:id="3"><w:p><w:r><w:t xml:space="preserve">Rewritten endnote</w:t></w:r></w:p></w:endnote>'
    );
    expect(written).not.toContain("Endnote body");
    // The footnotes part nothing touched goes out as the bytes it arrived as
    expect(
      bytesEqual(
        unzipSync(exported)["word/footnotes.xml"],
        unzipSync(bytes)["word/footnotes.xml"]
      )
    ).toBe(true);
  });

  it("creates an endnotes part beside a footnotes part that already exists", () => {
    const opened = importDocx(
      makeDeclaredDocx(
        `<w:p><w:r><w:t>Text</w:t></w:r>${reference("1")}` +
          '<w:r><w:endnoteReference w:id="1"/></w:r></w:p>'
      )
    );
    const edited = withStories(opened.doc, [
      [storyKey("footnote", "1"), storyFromText("A new footnote")],
      [storyKey("endnote", "1"), storyFromText("A new endnote")],
    ]);

    const exported = exportDocx(edited, opened.session);
    expect(writtenPart(exported, "word/_rels/document.xml.rels")).toMatch(
      new RegExp(
        `<Relationship (?=[^>]*Type="${ENDNOTES_REL}")(?=[^>]*Target="endnotes.xml")`
      )
    );
    expect(writtenPart(exported, "[Content_Types].xml")).toContain(
      `<Override PartName="/word/endnotes.xml" ContentType="${ENDNOTES_TYPE}"/>`
    );
    expect(writtenPart(exported, "word/endnotes.xml")).toMatch(
      /^<\?xml [^>]*\?><w:endnotes [^>]*><w:endnote w:type="separator" w:id="-1">.*<w:separator\/>.*<\/w:endnote><w:endnote w:type="continuationSeparator" w:id="0">.*<w:continuationSeparator\/>.*<\/w:endnote><w:endnote w:id="1">/
    );
    expect(writtenPart(exported, "word/footnotes.xml")).toContain(
      "A new footnote"
    );

    const again = importDocx(exported);
    expect(
      sameStory(
        storyNodeOf(again.doc, storyKey("endnote", "1")),
        storyOf("A new endnote")
      )
    ).toBe(true);
  });
});
