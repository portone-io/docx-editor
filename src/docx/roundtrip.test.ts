// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  documentXmlOf,
  drawingRun,
  exportErrorCode,
  fixtureNames,
  inlineDrawingXml,
  makeDocx,
  makeImageDocx,
  makeLinkedDocx,
  readFixture,
} from "../__testing__/docx";
import { createEditorState } from "../editor/createEditor";
import { parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import type { SessionStore } from "./session";

describe("round trip without editing", () => {
  it("there are fixtures", () => {
    expect(fixtureNames.length).toBeGreaterThan(0);
  });

  it.each(fixtureNames)(
    "%s: every part is byte identical to the original",
    (name) => {
      const bytes = readFixture(name);
      const { doc, session } = importDocx(bytes);
      const out = exportDocx(doc, session);

      const original = unzipSync(bytes);
      const exported = unzipSync(out);
      expect(Object.keys(exported)).toEqual(Object.keys(original));
      for (const key of Object.keys(original)) {
        if (key === session.mainPartPath) {
          expect(decode(exported[key])).toBe(decode(original[key]));
        }
        expect(bytesEqual(exported[key], original[key])).toBe(true);
      }
    }
  );
});

function namespaceDecls(xml: string): string[] {
  return xml.match(/xmlns(:[\w-]+)?="[^"]*"/g) ?? [];
}

function editFirstText(paragraph: PMNode): PMNode {
  const inline: PMNode[] = [];
  let edited = false;
  paragraph.forEach((child) => {
    if (!edited && child.isText && child.text) {
      inline.push(
        docxSchema.text("edited & <check> " + child.text, child.marks)
      );
      edited = true;
    } else {
      inline.push(child);
    }
  });
  if (!edited) throw new Error("no text to edit");
  return paragraph.copy(Fragment.from(inline));
}

describe("locality of an edit", () => {
  it.each(fixtureNames)(
    "%s: editing one paragraph creates no diff outside that paragraph",
    (name) => {
      const bytes = readFixture(name);
      const { doc, session } = importDocx(bytes);

      let targetIndex = -1;
      doc.forEach((child, _offset, i) => {
        if (
          targetIndex === -1 &&
          child.type.name === "paragraph" &&
          child.textContent.length > 0
        ) {
          targetIndex = i;
        }
      });
      expect(targetIndex).toBeGreaterThanOrEqual(0);

      const blocks: PMNode[] = [];
      doc.forEach((child, _offset, i) => {
        blocks.push(i === targetIndex ? editFirstText(child) : child);
      });
      const editedDoc = docxSchema.nodes.doc.create(null, blocks);
      const out = exportDocx(editedDoc, session);

      const exported = unzipSync(out);
      const documentXml = decode(exported[session.mainPartPath]);

      // The original fragments before and after the edited block are still there, byte for byte
      const head =
        session.documentPrefix +
        session.blocks
          .slice(0, targetIndex)
          .map((b) => b.xml)
          .join("");
      const tail =
        session.blocks
          .slice(targetIndex + 1)
          .map((b) => b.xml)
          .join("") + session.documentSuffix;
      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);

      // The regenerated fragment is valid XML and carries the edit
      const middle = documentXml.slice(
        head.length,
        documentXml.length - tail.length
      );
      expect(middle).toContain("edited &amp; &lt;check&gt;");
      // Editing does not splice in a namespace declaration that was not in the original
      expect(namespaceDecls(middle)).toEqual(
        namespaceDecls(session.blocks[targetIndex].xml)
      );
      const wrapped =
        '<w:wrap xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
        middle +
        "</w:wrap>";
      expect(() => parseXml(wrapped)).not.toThrow();

      // Every part other than the body is exactly as it was in the original
      const original = unzipSync(bytes);
      for (const key of Object.keys(original)) {
        if (key === session.mainPartPath) continue;
        expect(bytesEqual(exported[key], original[key])).toBe(true);
      }

      // Reopening the export reads the edit back
      const again = importDocx(out);
      expect(again.doc.textContent).toContain("edited & <check>");
    }
  );
});

function replaceAllText(doc: PMNode, text: string): PMNode {
  const blocks: PMNode[] = [];
  doc.forEach((block) => {
    const inline: PMNode[] = [];
    block.forEach((child) => {
      inline.push(child.isText ? docxSchema.text(text, child.marks) : child);
    });
    blocks.push(block.copy(Fragment.from(inline)));
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

describe("regenerating a paragraph", () => {
  it("a run with no text survives the edit too", () => {
    const emptyRun = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';
    const bytes = makeDocx(
      `<w:p>${emptyRun}<w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>`
    );
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(replaceAllText(doc, "edited text"), session);
    const documentXml = decode(unzipSync(out)["word/document.xml"]);

    expect(documentXml).toContain(emptyRun);
    expect(documentXml).toContain("edited text");
  });

  it("leaves a bookmark right where it was", () => {
    const bookmarkStart = '<w:bookmarkStart w:id="0" w:name="here"/>';
    const bookmarkEnd = '<w:bookmarkEnd w:id="0"/>';
    const bytes = makeDocx(
      `<w:p>${bookmarkStart}<w:r><w:t xml:space="preserve">source</w:t></w:r>${bookmarkEnd}</w:p>`
    );
    const { doc, session } = importDocx(bytes);
    const out = exportDocx(replaceAllText(doc, "edited text"), session);
    const documentXml = decode(unzipSync(out)["word/document.xml"]);

    expect(documentXml).toContain(`<w:p>${bookmarkStart}<w:r>`);
    expect(documentXml).toContain(`${bookmarkEnd}</w:p>`);
    expect(documentXml).toContain("edited text");
  });
});

/** The position just inside the first piece of text in the document */
function firstTextPos(doc: PMNode): number {
  let pos = -1;
  doc.descendants((node, at) => {
    if (pos === -1 && node.isText) pos = at + 1;
    return pos === -1;
  });
  if (pos === -1) throw new Error("no text");
  return pos;
}

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const cell = (text: string) => `<w:tc><w:p>${run(text)}</w:p></w:tc>`;

/** The document with the given edit applied through the editor, guards and all */
function afterEdit(
  doc: PMNode,
  edit: (tr: Transaction) => Transaction
): PMNode {
  const state = createEditorState(doc);
  return state.apply(edit(state.tr)).doc;
}

/** Every `w:id` value the XML carries */
function idValues(xml: string): string[] {
  return Array.from(xml.matchAll(/<w:id w:val="(\d+)"\/>/g)).map(
    (match) => match[1]
  );
}

const OPEN_SDT_PREFIX = '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>';

const LOCKED_SDT_PREFIX =
  "<w:sdt><w:sdtPr>" +
  '<w:id w:val="7"/><w:lock w:val="sdtContentLocked"/>' +
  "</w:sdtPr>";

const BOUND_SDT_PREFIX =
  "<w:sdt><w:sdtPr>" +
  '<w:id w:val="7"/><w:dataBinding w:xpath="/contract/date"/>' +
  "</w:sdtPr>";

const TBL_PR_EX =
  '<w:tblPrEx><w:tblBorders><w:top w:val="none"/></w:tblBorders></w:tblPrEx>';

/** A row carrying property exceptions, and a cell standing inside a control */
const WRAPPED_ROW_BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${TBL_PR_EX}<w:trPr><w:cantSplit w:val="0"/></w:trPr>` +
  `${OPEN_SDT_PREFIX}<w:sdtContent>${cell("value")}</w:sdtContent></w:sdt>` +
  cell("a") +
  "</w:tr></w:tbl>";

/** The same shape, with the control around the cell shutting it */
const LOCKED_CELL_BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${LOCKED_SDT_PREFIX}<w:sdtContent>${cell("final")}</w:sdtContent></w:sdt>` +
  cell("free") +
  "</w:tr></w:tbl>";

/** A paragraph with a locked control in the middle of its text */
const LOCKED_TEXT_BODY =
  "<w:p>" +
  run("signed on ") +
  `${LOCKED_SDT_PREFIX}<w:sdtContent>${run("2026-08-04")}</w:sdtContent></w:sdt>` +
  run(" final") +
  "</w:p>";

const twinControl = (text: string) =>
  `${OPEN_SDT_PREFIX}<w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;

/** Two controls written exactly alike, standing side by side */
const TWIN_CONTROLS_BODY = `<w:p>${twinControl("ab")}${twinControl("cd")}</w:p>`;

/** A control that a data binding points at */
const BOUND_CONTROL_BODY =
  "<w:p>" +
  `${BOUND_SDT_PREFIX}<w:sdtContent>${run("2026-08-04")}</w:sdtContent></w:sdt>` +
  "</w:p>";

const SEAL_DRAWING = drawingRun(inlineDrawingXml());

/** A control holding a piece of text and a picture */
const PICTURE_CONTROL_BODY =
  "<w:p>" +
  `${OPEN_SDT_PREFIX}<w:sdtContent>${run("seal")}${SEAL_DRAWING}</w:sdtContent></w:sdt>` +
  "</w:p>";

/** The three shapes a link arrives in, in one paragraph: external, anchor-only, and one carrying
 * the attributes we never read */
const LINK_BODY =
  "<w:p>" +
  run("see ") +
  `<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink>` +
  run(", ") +
  `<w:hyperlink w:anchor="chapter3">${run("chapter three")}</w:hyperlink>` +
  run(", and ") +
  '<w:hyperlink r:id="rId9" w:tooltip="Our terms" w:history="1" ' +
  `w:docLocation="top">${run("the top of it")}</w:hyperlink>` +
  "</w:p>";

const LINK_TARGETS = { rId9: "https://example.com/terms" };

const linkedDocx = (body: string) => makeLinkedDocx(body, LINK_TARGETS);

/** A control holding a link, which opens with both wrappers on the text */
const LINK_IN_CONTROL_BODY =
  "<w:p>" +
  `${OPEN_SDT_PREFIX}<w:sdtContent><w:hyperlink r:id="rId9">` +
  `${run("our terms")}</w:hyperlink></w:sdtContent></w:sdt>` +
  "</w:p>";

/** The markup Word leaves inside a control of its own accord */
const WORD_MARKUP_INSIDE =
  '<w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="1" w:name="signedOn"/>' +
  `${run("2026-08-04")}<w:bookmarkEnd w:id="1"/>`;

/** A locked control with that markup standing inside it */
const WORD_MARKUP_BODY =
  "<w:p>" +
  run("signed on ") +
  `${LOCKED_SDT_PREFIX}<w:sdtContent>${WORD_MARKUP_INSIDE}</w:sdtContent></w:sdt>` +
  "</w:p>";

/**
 * Every shape the import reads rather than preserves. Each of them is opened and written
 * straight back out, and the body has to come out as the very text it went in as
 */
const READ_STRUCTURES: ReadonlyArray<
  [string, string, (body: string) => Uint8Array]
> = [
  [
    "a row and a cell wrapped in markup we do not read",
    WRAPPED_ROW_BODY,
    makeDocx,
  ],
  ["a table cell Word locked", LOCKED_CELL_BODY, makeDocx],
  ["a paragraph holding a locked control", LOCKED_TEXT_BODY, makeDocx],
  ["two neighbouring controls written alike", TWIN_CONTROLS_BODY, makeDocx],
  ["a control a data binding points at", BOUND_CONTROL_BODY, makeDocx],
  ["a control holding a picture", PICTURE_CONTROL_BODY, makeImageDocx],
  ["a control Word left its own markup inside", WORD_MARKUP_BODY, makeDocx],
  ["a paragraph holding links of every shape", LINK_BODY, linkedDocx],
  ["a control holding a link", LINK_IN_CONTROL_BODY, linkedDocx],
];

describe.each(READ_STRUCTURES)("%s", (_name, body, make) => {
  it("goes back out byte identical as long as nothing is edited", () => {
    const bytes = make(body);
    const { doc, session } = importDocx(bytes);
    expect(documentXmlOf(doc, session)).toBe(
      decode(unzipSync(bytes)["word/document.xml"])
    );
  });
});

describe("a table whose row and cell carry wrappers we do not read", () => {
  it("opens as a table to edit rather than as a preserved block", () => {
    const { doc } = importDocx(makeDocx(WRAPPED_ROW_BODY));
    expect(doc.child(0).type.name).toBe("table");
  });

  it("both wrappers are still there after the cell text is edited", () => {
    const bytes = makeDocx(WRAPPED_ROW_BODY);
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("edit", firstTextPos(doc)));
    const documentXml = documentXmlOf(edited.doc, session);

    expect(documentXml).toContain("edit");
    expect(documentXml).toContain(TBL_PR_EX);
    expect(documentXml).toContain(`${OPEN_SDT_PREFIX}<w:sdtContent><w:tc>`);
    expect(documentXml).toContain("</w:tc></w:sdtContent></w:sdt>");
  });
});

/** The four values `w:lock` takes, each of which settles two clauses independently */
const LOCK_VALUES = [
  "unlocked",
  "sdtLocked",
  "contentLocked",
  "sdtContentLocked",
] as const;

const lockedPrefix = (val: string) =>
  "<w:sdt><w:sdtPr>" +
  `<w:id w:val="7"/><w:lock w:val="${val}"/>` +
  "</w:sdtPr>";

/**
 * The lock rides out inside the prefix the control is kept as, so what the two clauses were read
 * as never reaches the way back out. Both places a control can stand are checked over all four.
 */
describe.each(LOCK_VALUES)("a control whose lock reads %s", (val) => {
  const prefix = lockedPrefix(val);
  const BODIES = [
    "<w:p>" +
      run("signed on ") +
      `${prefix}<w:sdtContent>${run("2026-08-04")}</w:sdtContent></w:sdt>` +
      "</w:p>",
    "<w:tbl>" +
      '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
      `<w:tr>${prefix}<w:sdtContent>${cell("final")}</w:sdtContent></w:sdt>` +
      cell("free") +
      "</w:tr></w:tbl>",
  ];

  it.each(BODIES)("goes back out byte identical, nothing edited", (body) => {
    const bytes = makeDocx(body);
    const { doc, session } = importDocx(bytes);
    expect(documentXmlOf(doc, session)).toBe(
      decode(unzipSync(bytes)["word/document.xml"])
    );
  });
});

describe("a table whose cell Word locked", () => {
  /** The lock stands inside the wrapper the cell carries, so reading it changes nothing on the way out */
  it("opens as a table whose cell carries the lock", () => {
    const { doc } = importDocx(makeDocx(LOCKED_CELL_BODY));
    const row = doc.child(0).child(0);
    expect(row.child(0).attrs.sdtContentsLocked).toBe(true);
    expect(row.child(0).attrs.sdtPrefix).toBe(LOCKED_SDT_PREFIX);
    expect(row.child(1).attrs.sdtContentsLocked).toBe(false);
  });
});

describe("a paragraph holding a locked content control", () => {
  it("opens as a paragraph to edit rather than as a preserved block", () => {
    const { doc } = importDocx(makeDocx(LOCKED_TEXT_BODY));
    expect(doc.child(0).type.name).toBe("paragraph");
    expect(doc.child(0).textContent).toBe("signed on 2026-08-04 final");
  });

  it("the text inside the control comes in marked as locked", () => {
    const { doc } = importDocx(makeDocx(LOCKED_TEXT_BODY));
    const inside = doc.child(0).child(1);
    const mark = inside.marks.find((entry) => entry.type.name === "sdt");
    expect(inside.text).toBe("2026-08-04");
    expect(mark?.attrs.contentsLocked).toBe(true);
    expect(mark?.attrs.sdtPrefix).toBe(LOCKED_SDT_PREFIX);
  });

  it("the wrapper and its lock survive an edit to the text outside it", () => {
    const bytes = makeDocx(LOCKED_TEXT_BODY);
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("edit", firstTextPos(doc)));
    const documentXml = documentXmlOf(edited.doc, session);

    expect(documentXml).toContain("edit");
    expect(documentXml).toContain(
      `${LOCKED_SDT_PREFIX}<w:sdtContent>${run("2026-08-04")}</w:sdtContent></w:sdt>`
    );
    // The one control in the paragraph did not multiply on the way out
    expect(documentXml.match(/<w:sdt>/g)).toHaveLength(1);
  });
});

describe("two neighbouring controls written exactly alike", () => {
  it("open as two stretches of their own", () => {
    const { doc } = importDocx(makeDocx(TWIN_CONTROLS_BODY));
    const paragraph = doc.child(0);
    expect(paragraph.childCount).toBe(2);
    expect(paragraph.children.map((child) => child.text)).toEqual(["ab", "cd"]);
  });

  it("are still two controls once the text inside one of them is edited", () => {
    const { doc, session } = importDocx(makeDocx(TWIN_CONTROLS_BODY));
    const edited = afterEdit(doc, (tr) =>
      tr.insertText("edit", firstTextPos(doc))
    );
    const documentXml = documentXmlOf(edited, session);

    expect(documentXml).toContain(twinControl("aeditb"));
    expect(documentXml).toContain(twinControl("cd"));
    expect(documentXml.match(/<w:sdt>/g)).toHaveLength(2);
  });
});

describe("a control a paragraph split has left standing in two places", () => {
  /** The document with the control split down the middle, as pressing Enter inside it does */
  function split(): { doc: PMNode; session: SessionStore } {
    const { doc, session } = importDocx(makeDocx(BOUND_CONTROL_BODY));
    return {
      doc: afterEdit(doc, (tr) => tr.split(firstTextPos(doc) + 3)),
      session,
    };
  }

  it("writes a control in each paragraph, holding the text between them", () => {
    const { doc, session } = split();
    const documentXml = documentXmlOf(doc, session);

    expect(documentXml.match(/<w:sdt>/g)).toHaveLength(2);
    expect(documentXml).toContain(
      `<w:sdtContent>${run("2026")}</w:sdtContent>`
    );
    expect(documentXml).toContain(
      `<w:sdtContent>${run("-08-04")}</w:sdtContent>`
    );
  });

  it("keeps its id on the first of them and gives the copy one of its own", () => {
    const { doc, session } = split();
    const [first, second] = idValues(documentXmlOf(doc, session));

    expect(first).toBe("7");
    expect(second).not.toBe("7");
  });

  it("leaves the data binding with the first of them alone", () => {
    const { doc, session } = split();
    expect(documentXmlOf(doc, session).match(/<w:dataBinding /g)).toHaveLength(
      1
    );
  });
});

describe("a control holding a picture", () => {
  it("opens with the picture inside the control, editable", () => {
    const { doc } = importDocx(makeImageDocx(PICTURE_CONTROL_BODY));
    const paragraph = doc.child(0);
    expect(paragraph.children.map((child) => child.type.name)).toEqual([
      "text",
      "image",
    ]);
    const mark = paragraph
      .child(1)
      .marks.find((entry) => entry.type.name === "sdt");
    expect(mark?.attrs.sdtPrefix).toBe(OPEN_SDT_PREFIX);
  });

  it("keeps the picture inside the control after the text beside it is edited", () => {
    const { doc, session } = importDocx(makeImageDocx(PICTURE_CONTROL_BODY));
    const edited = afterEdit(doc, (tr) =>
      tr.insertText("edit", firstTextPos(doc) + 1)
    );
    const documentXml = documentXmlOf(edited, session);

    expect(documentXml).toContain("edit");
    expect(documentXml).toContain(`${SEAL_DRAWING}</w:sdtContent></w:sdt>`);
    expect(documentXml.match(/<w:sdt>/g)).toHaveLength(1);
  });
});

describe("a control Word has left its own markup inside", () => {
  it("opens as a paragraph to edit rather than as a preserved block", () => {
    const { doc } = importDocx(makeDocx(WORD_MARKUP_BODY));
    expect(doc.child(0).type.name).toBe("paragraph");
    expect(doc.child(0).textContent).toBe("signed on 2026-08-04");
  });

  it("keeps the lock on everything standing inside it", () => {
    const { doc } = importDocx(makeDocx(WORD_MARKUP_BODY));
    const locked = doc
      .child(0)
      .children.map((child) =>
        child.marks.some(
          (entry) =>
            entry.type.name === "sdt" && entry.attrs.contentsLocked === true
        )
      );
    // The run ahead of the control is open; the markup and the text within it are shut
    expect(locked).toEqual([false, true, true, true, true]);
  });

  it("keeps that markup where it stood once the text outside is edited", () => {
    const { doc, session } = importDocx(makeDocx(WORD_MARKUP_BODY));
    const edited = afterEdit(doc, (tr) =>
      tr.insertText("edit", firstTextPos(doc))
    );
    const documentXml = documentXmlOf(edited, session);

    expect(documentXml).toContain("edit");
    expect(documentXml).toContain(
      `${LOCKED_SDT_PREFIX}<w:sdtContent>${WORD_MARKUP_INSIDE}</w:sdtContent></w:sdt>`
    );
  });
});

describe("a paragraph holding links", () => {
  const LINK_XML = /<w:hyperlink[^>]*>/g;

  it("opens as editable text wearing the address of each link", () => {
    const { doc } = importDocx(linkedDocx(LINK_BODY));
    const paragraph = doc.child(0);
    expect(paragraph.textContent).toBe(
      "see our terms, chapter three, and the top of it"
    );
    const addresses = paragraph.children.map(
      (child) =>
        child.marks.find((mark) => mark.type.name === "link")?.attrs.href ??
        null
    );
    // The anchor-only link is editable text like the others, and offers no address
    expect(addresses).toEqual([
      null,
      "https://example.com/terms",
      null,
      null,
      null,
      "https://example.com/terms",
    ]);
  });

  it("keeps every wrapper as it came once the text outside them is edited", () => {
    const bytes = linkedDocx(LINK_BODY);
    const { doc, session } = importDocx(bytes);
    const edited = afterEdit(doc, (tr) => tr.insertText("edit", 1));
    const documentXml = documentXmlOf(edited, session);

    expect(documentXml).toContain("edit");
    expect(documentXml.match(LINK_XML)).toEqual(
      decode(unzipSync(bytes)["word/document.xml"]).match(LINK_XML)
    );
  });

  it("adds no relationship for a link that was already there", () => {
    const bytes = linkedDocx(LINK_BODY);
    const { doc, session } = importDocx(bytes);
    const edited = afterEdit(doc, (tr) => tr.insertText("edit", 1));
    const out = unzipSync(exportDocx(edited, session));
    expect(decode(out["word/_rels/document.xml.rels"])).toBe(
      decode(unzipSync(bytes)["word/_rels/document.xml.rels"])
    );
  });

  it("writes a wrapper of its own in each paragraph a split left the link standing in", () => {
    const bytes = linkedDocx(
      `<w:p><w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink></w:p>`
    );
    const { doc, session } = importDocx(bytes);
    const split = afterEdit(doc, (tr) => tr.split(firstTextPos(doc) + 3));
    const documentXml = documentXmlOf(split, session);

    // Both halves point at the one relationship, which nothing in the standard makes single-use
    expect(documentXml.match(LINK_XML)).toEqual([
      '<w:hyperlink r:id="rId9">',
      '<w:hyperlink r:id="rId9">',
    ]);
    expect(documentXml).toContain(`${run("our ")}</w:hyperlink></w:p>`);
  });
});

describe("a preserved block that lost its original", () => {
  it("refuses to export", () => {
    const { session } = importDocx(makeDocx("<w:p/>"));
    const doc = docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.rawBlock.create({ name: "w:tbl" }),
    ]);
    expect(exportErrorCode(() => exportDocx(doc, session))).toBe(
      "lost-original"
    );
  });
});

describe("the section setup at the end of the body", () => {
  const PAGE_SETUP = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
  const bytes = makeDocx(
    `<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>${PAGE_SETUP}`
  );

  it("does not open as a block to edit", () => {
    const { doc } = importDocx(bytes);
    expect(doc.childCount).toBe(1);
    expect(doc.child(0).type.name).toBe("paragraph");
  });

  /**
   * Left as a document node, it used to disappear on select-all + delete.
   * Whatever was written afterwards then went out with no page setup, and Word opened it on the
   * default paper.
   */
  it("survives in the file even after select all, delete and retype", () => {
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const wiped = state.apply(
      state.tr.delete(0, state.doc.content.size).insertText("rewritten text", 1)
    );
    const documentXml = decode(
      unzipSync(exportDocx(wiped.doc, session))["word/document.xml"]
    );

    expect(documentXml).toContain("rewritten text");
    expect(documentXml).toContain(PAGE_SETUP);
  });
});
