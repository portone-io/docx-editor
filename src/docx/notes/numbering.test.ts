// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import { decode, makeNotesDocx } from "../../__testing__/docx";
import { exportDocx } from "../exportDocx";
import { importDocx } from "../importDocx";
import { withSectionBreak } from "../sections";
import { withNoteLabels } from "./numbering";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SETTINGS_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const A4 = '<w:pgSz w:w="11906" w:h="16838"/>';

const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const footnote = (id: string, attrs = "") =>
  `<w:r><w:footnoteReference w:id="${id}"${attrs}/></w:r>`;
const endnote = (id: string) => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;
const paragraph = (...runs: string[]) => `<w:p>${runs.join("")}</w:p>`;

/** A paragraph that ends a section laying down these properties */
const sectionEnd = (sectPr: string, ...runs: string[]) =>
  `<w:p><w:pPr><w:sectPr>${sectPr}${A4}</w:sectPr></w:pPr>${runs.join("")}</w:p>`;

/** The notes package, with a settings part holding these children */
function withSettings(body: string, settings: string): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeNotesDocx(body));
  const rels = "word/_rels/document.xml.rels";
  parts[rels] = encoder.encode(
    decode(parts[rels]).replace(
      "</Relationships>",
      `<Relationship Id="rId6" Target="settings.xml" Type="${SETTINGS_REL}"/></Relationships>`
    )
  );
  parts["word/settings.xml"] = encoder.encode(
    `<w:settings xmlns:w="${W_NS}">${settings}</w:settings>`
  );
  return zipSync(parts);
}

function opened(body: string, settings = ""): PMNode {
  return importDocx(withSettings(body, settings)).doc;
}

/** Each reference of the body as its kind and the label it is drawn with, in document order */
function labels(doc: PMNode): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "noteReference") {
      found.push(`${node.attrs.kind} ${node.attrs.label}`);
    }
    return true;
  });
  return found;
}

describe("noteLabelsIn", () => {
  it("labels footnotes in first-reference order from one", () => {
    const doc = opened(
      paragraph(footnote("5"), text("then"), footnote("2")) +
        paragraph(footnote("9"))
    );

    expect(labels(doc)).toEqual(["footnote 1", "footnote 2", "footnote 3"]);
  });

  it("spells endnote labels in decimal when the settings name no format", () => {
    const { doc, session } = importDocx(
      withSettings(
        paragraph(footnote("2"), endnote("3"), endnote("4")),
        '<w:footnotePr><w:numFmt w:val="upperLetter"/></w:footnotePr>'
      )
    );

    expect(session.noteNumbering.endnote.format).toBe("decimal");
    expect(labels(doc)).toEqual(["footnote A", "endnote 1", "endnote 2"]);
  });

  it("starts counting at the number the settings name", () => {
    const doc = opened(
      paragraph(footnote("2"), endnote("3"), footnote("4")),
      '<w:footnotePr><w:numStart w:val="4"/></w:footnotePr>' +
        '<w:endnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/></w:endnotePr>'
    );

    expect(labels(doc)).toEqual(["footnote 4", "endnote iii", "footnote 5"]);
  });

  it("counts again from the start at each section when the restart is eachSect", () => {
    const doc = opened(
      paragraph(footnote("2"), endnote("3")) +
        sectionEnd("", footnote("4")) +
        paragraph(footnote("5"), endnote("6")),
      '<w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr>'
    );

    expect(labels(doc)).toEqual([
      "footnote 1",
      "endnote 1",
      "footnote 2",
      "footnote 1",
      "endnote 2",
    ]);
  });

  it("takes a section's own footnote properties over the settings", () => {
    const doc = opened(
      sectionEnd(
        '<w:footnotePr><w:numFmt w:val="upperRoman"/></w:footnotePr>',
        footnote("2"),
        footnote("3")
      ) + paragraph(footnote("4")),
      '<w:footnotePr><w:numFmt w:val="lowerLetter"/></w:footnotePr>'
    );

    expect(labels(doc)).toEqual(["footnote I", "footnote II", "footnote c"]);
  });

  it("gives a repeated reference the label of its first", () => {
    const doc = opened(
      paragraph(footnote("2"), footnote("3"), text("again"), footnote("2"))
    );

    expect(labels(doc)).toEqual(["footnote 1", "footnote 2", "footnote 1"]);
  });

  it("gives a reference repeated after a restart the number of its own section", () => {
    const doc = opened(
      sectionEnd("", footnote("2"), footnote("3")) +
        paragraph(footnote("2"), footnote("4")),
      '<w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr>'
    );

    expect(labels(doc)).toEqual([
      "footnote 1",
      "footnote 2",
      "footnote 1",
      "footnote 2",
    ]);
  });

  it("gives a reference repeated in a later section its first label where the count runs on", () => {
    const doc = opened(
      sectionEnd("", footnote("2"), footnote("3")) +
        paragraph(footnote("2"), footnote("4"))
    );

    expect(labels(doc)).toEqual([
      "footnote 1",
      "footnote 2",
      "footnote 1",
      "footnote 3",
    ]);
  });

  it("skips a reference whose custom mark follows", () => {
    const doc = opened(
      paragraph(
        footnote("2"),
        footnote("3", ' w:customMarkFollows="1"'),
        text("*"),
        footnote("4")
      )
    );

    expect(labels(doc)).toEqual(["footnote 1", "footnote ", "footnote 2"]);
  });

  it("counts a restart on each page straight through", () => {
    const doc = opened(
      sectionEnd(
        "",
        footnote("2"),
        text("page"),
        '<w:r><w:br w:type="page"/></w:r>'
      ) + paragraph(footnote("3")),
      '<w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>'
    );

    expect(labels(doc)).toEqual(["footnote 1", "footnote 2"]);
  });

  it("gives no number to a reference that calls a separator entry", () => {
    const doc = opened(paragraph(footnote("-1"), footnote("2")));

    expect(labels(doc)).toEqual(["footnote ?", "footnote 1"]);
  });

  it("labels a document reached by edits as the file it exports opens", () => {
    const { doc, session } = importDocx(
      withSettings(
        paragraph(footnote("2"), endnote("3")) +
          paragraph(text("second"), footnote("5")) +
          paragraph(footnote("6", ' w:customMarkFollows="1"'), text("*")) +
          paragraph(footnote("7"), endnote("8")),
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numRestart w:val="eachSect"/></w:footnotePr>' +
          '<w:endnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/></w:endnotePr>'
      )
    );
    const edit = new Transform(doc).delete(1, 2);
    const second = edit.doc.child(1);
    const pPr: unknown = second.attrs.pPr;
    edit.setNodeMarkup(edit.doc.child(0).nodeSize, undefined, {
      ...second.attrs,
      pPr: withSectionBreak(
        typeof pPr === "string" ? pPr : null,
        `<w:sectPr>${A4}</w:sectPr>`
      ),
    });
    const edited = withNoteLabels(
      edit.doc,
      session.noteNumbering,
      session.specialNotes
    );
    const reopened = importDocx(exportDocx(edited, session)).doc;

    expect(labels(edited)).toEqual([
      "endnote III",
      "footnote a",
      "footnote ",
      "footnote a",
      "endnote IV",
    ]);
    expect(labels(reopened)).toEqual(labels(edited));
  });
});
