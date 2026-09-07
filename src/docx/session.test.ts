// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { makeDocx, makeStyledNumberedDocx } from "../__testing__/docx";
import { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import {
  BODY_STORY_KEY,
  blockKey,
  documentNumbering,
  originalBlock,
  type SessionIdentity,
  splitBlockKey,
} from "./session";

/** A document whose second block is one the editor does not model, so it is kept as it came */
const bodyOf = (name: string) =>
  `<w:p><w:r><w:t>${name} text</w:t></w:r></w:p>` +
  `<w:customXml w:uri="urn:${name}" w:element="from-${name}"/>`;

function open(name: string) {
  return importDocx(makeDocx(bodyOf(name)));
}

/** A preserved block standing where `key` says it came from */
function preserved(key: string) {
  return docxSchema.nodes.docxRaw.create({
    srcId: key,
    name: "w:customXml",
  });
}

describe("the original a block points at", () => {
  it("a block of another session is not an original of this one", () => {
    const { session } = open("alpha");
    const elsewhere: SessionIdentity = { sessionId: "another-document" };

    expect(
      originalBlock(preserved(blockKey(session, BODY_STORY_KEY, 1)), session)
        ?.xml
    ).toBe('<w:customXml w:uri="urn:alpha" w:element="from-alpha"/>');
    expect(
      originalBlock(preserved(blockKey(elsewhere, BODY_STORY_KEY, 1)), session)
    ).toBeUndefined();
  });

  it("a block key of another story is not an original of the body", () => {
    const { session } = open("alpha");

    expect(
      originalBlock(preserved(blockKey(session, "comment:4", 1)), session)
    ).toBeUndefined();
  });
});

describe("a block key", () => {
  it.each([
    ["an empty story", "opened::0"],
    ["an index that loses precision", "opened:body:9007199254740993"],
    ["an index that overflows", `opened:body:${"9".repeat(400)}`],
  ])("refuses %s", (_description, key) => {
    expect(splitBlockKey(key)).toBeNull();
  });

  it.each([0, Number.MAX_SAFE_INTEGER])("keeps the exact index %s", (index) => {
    const session = { sessionId: "opened" };
    expect(splitBlockKey(blockKey(session, "comment:4", index))).toEqual({
      sessionId: "opened",
      storyKey: "comment:4",
      index,
    });
  });

  it("splitBlockKey reads a story key that itself holds a colon", () => {
    expect(splitBlockKey("d3-1a2b:comment:4:7")).toEqual({
      sessionId: "d3-1a2b",
      storyKey: "comment:4",
      index: 7,
    });
  });

  it("is not read out of a string that is not one", () => {
    expect(splitBlockKey("1")).toBeNull();
    expect(splitBlockKey("d3-1a2b:body")).toBeNull();
    expect(splitBlockKey("d3-1a2b:body:last")).toBeNull();
  });
});

describe("exporting a block from elsewhere", () => {
  it("a preserved block of another session refuses export as lost-original naming the other document", () => {
    const alpha = open("alpha");
    const beta = open("beta");
    const moved = docxSchema.nodes.doc.create(null, [
      beta.doc.child(0),
      alpha.doc.child(1),
    ]);

    try {
      exportDocx(moved, beta.session);
      throw new Error("the export was expected to be refused");
    } catch (error) {
      expect(error).toBeInstanceOf(DocxExportError);
      const refusal = error as DocxExportError;
      expect(refusal.code).toBe("lost-original");
      expect(refusal.message).toContain(alpha.session.sessionId);
    }
  });
});

describe("opening a document twice", () => {
  it("the session id differs between two imports of the same bytes", () => {
    const bytes = makeDocx(bodyOf("alpha"));

    expect(importDocx(bytes).session.sessionId).not.toBe(
      importDocx(bytes).session.sessionId
    );
  });
});

/**
 * A consumer asks an open document for its lists through a reader of its own, which has to read
 * the numbering part the way opening the document read it: the numbering styles of styles.xml
 * resolved, and the formatting each level puts on its number read in.
 */
describe("the lists an open document hands back", () => {
  const NUMBERING_STYLE =
    '<w:style w:type="numbering" w:styleId="Clauses">' +
    '<w:pPr><w:numPr><w:numId w:val="6"/></w:numPr></w:pPr></w:style>';

  const W_NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  const LINKED_NUMBERING =
    `<w:numbering ${W_NS}>` +
    '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="Clauses"/></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">' +
    '<w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/>' +
    "<w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum>" +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="6"><w:abstractNumId w:val="1"/></w:num></w:numbering>';

  it("follows the numbering styles and reads what each level dresses its number in", () => {
    const { session } = importDocx(
      makeStyledNumberedDocx(
        "<w:p><w:r><w:t>text</w:t></w:r></w:p>",
        NUMBERING_STYLE,
        LINKED_NUMBERING
      )
    );

    const level = documentNumbering(session).lists.get(1)?.levels.get(0);
    expect(level?.format).toBe("upperRoman");
    expect(level?.run).toEqual({ bold: true });
  });
});
