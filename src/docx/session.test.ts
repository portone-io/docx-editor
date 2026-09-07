// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import {
  BODY_STORY_KEY,
  blockKey,
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
