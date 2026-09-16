// @vitest-environment jsdom

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  decode,
  documentXmlOf,
  makeDocx,
  makeLinkedDocx,
  readFixture,
  withBlocks,
} from "../__testing__/docx";
import { rangeOfText, runCommand, select } from "../__testing__/editing";
import { undo } from "../editor/commands";
import { setLink } from "../editor/commands/linkCommands";
import { createEditorState } from "../editor/createEditor";
import { parseXml, R_NS } from "../ooxml/xml";
import { wrapperMarks } from "../schema/wrappers";
import { firstBlockIndex, withEditedBlock } from "./__testing__/blockEdits";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import type { LinkTargets } from "./hyperlink";
import { importDocx } from "./importDocx";
import { buildParagraph, NO_IMPORT_SOURCES } from "./importParagraph";
import { serializeParagraph } from "./serializeParagraph";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const TERMS = "https://example.com/terms";
const LINKS: LinkTargets = new Map([["rId9", TERMS]]);

/** The relationships an export hands out: the one address, on the id it came in on */
const REFS: ExportRefs = {
  ...NO_EXPORT_REFS,
  links: { relIdOf: (href) => (href === TERMS ? "rId9" : undefined) },
};

function open(xml: string): PMNode {
  const wrapped = parseXml(`<w:wrap ${W_NS} xmlns:r="${R_NS}">${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return buildParagraph(el, null, { ...NO_IMPORT_SOURCES, links: LINKS });
}

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A control, written the way a file writes one: an id, then an optional lock */
const control = (inner: string, { id = 7, lock = "" } = {}) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/>` +
  (lock === "" ? "" : `<w:lock w:val="${lock}"/>`) +
  `</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const link = (inner: string) =>
  `<w:hyperlink r:id="rId9">${inner}</w:hyperlink>`;

/** The one inline node a paragraph holding nothing else opened as */
function only(node: PMNode): PMNode {
  expect(node.type.name).toBe("paragraph");
  expect(node.childCount).toBe(1);
  return node.child(0);
}

/** What every wrapper the node stands inside is, outermost first */
function nesting(node: PMNode): string[] {
  return wrapperMarks(node).map((mark) => mark.type.name);
}

describe("a wrapper holding a wrapper", () => {
  it("a hyperlink holding a control round-trips with the link outside", () => {
    const xml = `<w:p>${link(control(run("terms")))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["link", "sdt"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("a control holding a link round-trips with the control outside", () => {
    const xml = `<w:p>${control(link(run("terms")))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["sdt", "link"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("two nested controls keep their order and both ids", () => {
    const xml = `<w:p>${control(control(run("terms"), { id: 8 }))}</w:p>`;
    const node = open(xml);
    const marks = wrapperMarks(only(node));

    expect(marks.map((mark) => mark.attrs.sdtPrefix)).toEqual([
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>',
      '<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr>',
    ]);
    // Two controls written alike would still be two, so each carries its own number
    expect(marks.map((mark) => mark.attrs.key)).toEqual([0, 1]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("three wrappers deep still rebuild in file order", () => {
    const xml = `<w:p>${control(link(control(run("terms"), { id: 8 })))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["sdt", "link", "sdt"]);
    expect(wrapperMarks(only(node)).map((mark) => mark.attrs.depth)).toEqual([
      0, 1, 2,
    ]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("keeps the text beside a nested wrapper outside it", () => {
    const xml = `<w:p>${run("a")}${link(`${run("b")}${control(run("c"))}${run("d")}`)}${run("e")}</w:p>`;
    const node = open(xml);

    expect(node.textContent).toBe("abcde");
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });
});

/**
 * `w:customXml` wraps inline content exactly as a control does, and no kind is registered for it,
 * so the paragraph around it stays editable and the wrapper is kept whole where it stood
 * (`docx/importPolicy`). It is not the whole paragraph that is stood down.
 */
describe("a wrapper kind nobody registered", () => {
  it("is kept whole inside a paragraph that stays editable", () => {
    const xml = `<w:p>${run("a")}<w:customXml w:element="x">${run("b")}</w:customXml></w:p>`;
    const node = open(xml);

    expect(node.type.name).toBe("paragraph");
    expect(node.children.map((child) => child.type.name)).toEqual([
      "text",
      "rawInline",
    ]);
    expect(node.child(1).attrs.element).toBe("customXml");
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("keeps the wrappers around it, so they still close where they did", () => {
    const xml = `<w:p>${link(`<w:customXml w:element="x">${run("b")}</w:customXml>`)}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["link"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });
});

/**
 * `EG_PContent` admits it and Word never writes it. A link mark cannot record a second link around
 * it, so the inner one is kept whole where it stood, wearing the link around it, exactly as it was
 * before any of this nested.
 */
describe("a hyperlink inside a hyperlink", () => {
  const OTHER = "https://example.com/other";

  it("keeps the inner one whole and round-trips as it came", () => {
    const xml = `<w:p>${link(`<w:hyperlink r:id="rId8">${run("terms")}</w:hyperlink>`)}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["link"]);
    expect(only(node).type.name).toBe("rawInline");
    expect(only(node).attrs.element).toBe("hyperlink");
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("round-trips byte identical after an edit elsewhere in the paragraph", () => {
    const body =
      `<w:p>${run("see ")}` +
      `<w:hyperlink r:id="rId8"><w:hyperlink r:id="rId9">${run("terms")}</w:hyperlink></w:hyperlink>` +
      "</w:p>";
    const bytes = makeLinkedDocx(body, { rId9: TERMS, rId8: OTHER });
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const typed = state.apply(state.tr.insertText("X", 2));
    const written = documentXmlOf(typed.doc, session);

    // The inner link is a fragment kept whole, so its text is not the paragraph's to edit
    expect(typed.doc.child(0).textContent).toBe("sXee ");
    // The edit reached the block, so what follows is the writer's work rather than the original
    // bytes handed back
    expect(written).toContain("sXee ");
    // Both wrappers still stand: the inner one does not push the outer off the content it wraps
    expect(written).toContain(
      `<w:hyperlink r:id="rId8"><w:hyperlink r:id="rId9">${run("terms")}</w:hyperlink></w:hyperlink>`
    );
    expect(written.match(/<w:hyperlink/g)).toHaveLength(2);
  });
});

describe("a document holding a hyperlink that holds a control", () => {
  const BODY = `<w:p>${link(control(run("terms")))}</w:p>`;

  it("opens as an editable paragraph and exports byte identical", () => {
    const bytes = makeLinkedDocx(BODY, { rId9: TERMS });
    const { doc, session } = importDocx(bytes);

    expect(doc.child(0).type.name).toBe("paragraph");
    expect(documentXmlOf(doc, session)).toBe(
      decode(unzipSync(bytes)["word/document.xml"])
    );
  });

  it("typing inside the control stays inside both wrappers", () => {
    const bytes = makeLinkedDocx(BODY, { rId9: TERMS });
    const state = createEditorState(importDocx(bytes).doc);
    const { from } = rangeOfText(state.doc, "terms");
    const typed = state.apply(state.tr.insertText("X", from + 2));
    const paragraph = typed.doc.child(0);

    expect(paragraph.textContent).toBe("teXrms");
    expect(nesting(only(paragraph))).toEqual(["link", "sdt"]);
    expect(serializeParagraph(paragraph, REFS)).toContain("teXrms");
  });

  it("a link made in the editor inside a control lands inside it", () => {
    const bytes = makeLinkedDocx(
      `<w:p>${control(run("read the terms"))}</w:p>`,
      {
        rId9: TERMS,
      }
    );
    const { doc } = importDocx(bytes);
    const state = createEditorState(doc);
    const { from, to } = rangeOfText(state.doc, "terms");
    const linked = runCommand(select(state, from, to), setLink(TERMS));

    const marks = wrapperMarks(linked.doc.resolve(from + 1).parent.child(1));
    expect(marks.map((mark) => mark.type.name)).toEqual(["sdt", "link"]);
    expect(marks.map((mark) => mark.attrs.depth)).toEqual([0, 1]);
    // One control still, with the link written inside the one `w:sdtContent`
    const written = serializeParagraph(linked.doc.child(0), REFS);
    expect(written.match(/<w:sdt[ >]/g)).toHaveLength(1);
    expect(written).toContain(
      `<w:hyperlink r:id="rId9">${run("terms")}</w:hyperlink></w:sdtContent>`
    );
  });

  /**
   * The pieces a link goes on are one per inline node, so a link laid over two runs gets a mark
   * each. They carry the same address at the same depth, and the writer groups by what the marks
   * say rather than by which object they are, so one hyperlink comes out.
   */
  it("a link made across two differently formatted runs goes out as one hyperlink", () => {
    const bytes = makeLinkedDocx(
      `<w:p>${run("our ")}<w:r><w:rPr><w:b/></w:rPr><w:t>terms</w:t></w:r></w:p>`,
      { rId9: TERMS }
    );
    const state = createEditorState(importDocx(bytes).doc);
    const { from } = rangeOfText(state.doc, "our ");
    const { to } = rangeOfText(state.doc, "terms");
    const linked = runCommand(select(state, from, to), setLink(TERMS));
    const written = serializeParagraph(linked.doc.child(0), REFS);

    expect(written.match(/<w:hyperlink/g)).toHaveLength(1);
    expect(written).toContain("</w:hyperlink></w:p>");
  });

  it("undoing that link leaves the control exactly as it was", () => {
    const bytes = makeLinkedDocx(
      `<w:p>${control(run("read the terms"))}</w:p>`,
      { rId9: TERMS }
    );
    const state = createEditorState(importDocx(bytes).doc);
    const before = serializeParagraph(state.doc.child(0), REFS);
    const { from, to } = rangeOfText(state.doc, "terms");
    const linked = runCommand(select(state, from, to), setLink(TERMS));

    expect(
      serializeParagraph(runCommand(linked, undo).doc.child(0), REFS)
    ).toBe(before);
  });
});

/**
 * The cases below are built from small bodies rather than read off the fixture, so what each one
 * proves stands in the test itself; `content-controls.docx` is what holds them in one package and
 * is read for the byte identity and the edited round trip.
 */
const FIXTURE = "content-controls.docx";

const P = (inline: string) => `<w:p>${inline}</w:p>`;

/** The type of every inline node of the one paragraph a body holds */
function inlineTypes(body: string): string[] {
  const doc = importDocx(makeDocx(body)).doc;
  const names: string[] = [];
  doc.child(0).forEach((child) => {
    names.push(child.type.name);
  });
  return names;
}

/** The one control a paragraph holds, which every case here writes as its second inline node */
function emptyControl(body: string): PMNode {
  const node = importDocx(makeDocx(body)).doc.child(0).child(1);
  expect(node.type.name).toBe("sdtEmptyInline");
  return node;
}

/**
 * Which node a wrapper holding nothing opens as is the registry's answer, one entry at a time, so a
 * kind with no node of its own keeps the element whole where it stood.
 */
describe("a wrapper a paragraph holds with nothing inside it", () => {
  it("keeps a hyperlink of that shape whole where it stood", () => {
    expect(
      inlineTypes(
        P(
          run("before") +
            '<w:hyperlink r:id="rId99"></w:hyperlink>' +
            run("after")
        )
      )
    ).toEqual(["text", "rawInline", "text"]);
  });
});

describe("a control a paragraph holds with nothing inside it", () => {
  it("is read as the node that draws nothing rather than as a preserved chip", () => {
    expect(
      inlineTypes(P(run("before") + control("", { id: 1 }) + run("after")))
    ).toEqual(["text", "sdtEmptyInline", "text"]);
  });

  it("carries what the mark would carry: the opening tag and a key", () => {
    const node = emptyControl(
      P(run("before") + control("", { id: 42, lock: "sdtLocked" }))
    );

    expect(node.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="42"/><w:lock w:val="sdtLocked"/></w:sdtPr>'
    );
    expect(node.attrs).toMatchObject({
      contentsLocked: false,
      deletionLocked: true,
      key: expect.any(Number),
    });
  });

  /**
   * `CT_Sdt` writes `w:sdtContent` `minOccurs="0"`, so a control may state that what it stood
   * around is not there by leaving the element out, and §17.5.2.34 makes that element a cache of
   * contents rather than the statement itself. The two shapes are therefore read as one node.
   */
  it("is read the same when the file wrote no content element at all", () => {
    const node = emptyControl(
      P(
        run("before") +
          '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
          '<w:lock w:val="sdtContentLocked"/></w:sdtPr></w:sdt>'
      )
    );

    expect(node.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
        '<w:lock w:val="sdtContentLocked"/></w:sdtPr>'
    );
    expect(node.attrs).toMatchObject({
      contentsLocked: true,
      deletionLocked: true,
    });
  });

  it("carries a w:sdtEndPr written with no content element in the prefix", () => {
    const node = emptyControl(
      P(
        run("before") +
          '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>' +
          "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr></w:sdt>"
      )
    );

    expect(node.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>' +
        "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>"
    );
  });

  it("is still kept whole where it carries no w:sdtPr to read", () => {
    expect(inlineTypes(P(run("before") + "<w:sdt/>"))).toEqual([
      "text",
      "rawInline",
    ]);
  });

  it("counts one control apart from the next, as the mark does", () => {
    const doc = importDocx(
      makeDocx(
        P(
          run("before") +
            control("", { id: 1 }) +
            control(run("held"), { id: 2 })
        )
      )
    ).doc;

    expect(doc.child(0).child(1).attrs.key).not.toBe(
      doc.child(0).child(2).marks[0]?.attrs.key
    );
  });

  it("wears a hyperlink's mark as readily as a control's", () => {
    const doc = importDocx(
      makeDocx(
        P(
          run("before") +
            `<w:hyperlink w:anchor="here">${run("linked")}${control("", { id: 2 })}` +
            "</w:hyperlink>"
        )
      )
    ).doc;
    const node = doc.child(0).child(2);

    expect(node.type.name).toBe("sdtEmptyInline");
    expect(node.marks.map((mark) => mark.type.name)).toEqual(["link"]);
  });

  it("wears the marks of the wrappers it stands inside", () => {
    const doc = importDocx(
      makeDocx(
        P(
          run("before") +
            control(run("held") + control("", { id: 2 }), { id: 1 })
        )
      )
    ).doc;
    const node = doc.child(0).child(2);

    expect(node.type.name).toBe("sdtEmptyInline");
    expect(node.marks.map((mark) => mark.type.name)).toEqual(["sdt"]);
    expect(node.marks[0]?.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr>'
    );
  });
});

describe("a control a paragraph holds with nothing inside it, written back", () => {
  it("goes back out as the bytes it arrived as while an edit stands beside it", () => {
    const { doc, session } = importDocx(readFixture(FIXTURE));
    const index = firstBlockIndex(doc, "paragraph");

    const out = documentXmlOf(withEditedBlock(doc, index, "Edited"), session);

    expect(out).toContain(
      '<w:sdt><w:sdtPr><w:alias w:val="Signing date"/><w:id w:val="416"/>' +
        "</w:sdtPr><w:sdtContent></w:sdtContent></w:sdt>"
    );
    expect(documentXmlOf(doc, session)).toBe(
      decode(unzipSync(readFixture(FIXTURE))[session.mainPartPath])
    );
  });

  /**
   * A paragraph nobody rewrote never reaches the writer, so the paragraph is edited here: the
   * control is then built from its attrs, content tag and all, and whichever of the two shapes it
   * arrived as it goes out as the one the writer knows how to take back apart.
   */
  it("writes the content tag itself for a paragraph the writer rebuilt", () => {
    const original = '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr></w:sdt>';
    const { doc, session } = importDocx(
      makeDocx(P(run("before") + original + run("after")))
    );

    const out = documentXmlOf(withEditedBlock(doc, 0, "Edited"), session);

    expect(out).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr>' +
        "<w:sdtContent></w:sdtContent></w:sdt>"
    );
    expect(out).not.toContain(original);
  });

  it("closes a hyperlink it stands inside around it again", () => {
    const { doc, session } = importDocx(
      makeDocx(
        P(
          run("before") +
            `<w:hyperlink w:anchor="here">${run("linked")}${control("", { id: 2 })}` +
            "</w:hyperlink>"
        )
      )
    );

    const out = documentXmlOf(withEditedBlock(doc, 0, "Edited"), session);

    expect(out).toContain(
      '<w:hyperlink w:anchor="here">' +
        '<w:r><w:t xml:space="preserve">linked</w:t></w:r>' +
        '<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>' +
        "</w:sdtContent></w:sdt></w:hyperlink>"
    );
  });

  it("closes the wrapper it stands inside around it again", () => {
    const { doc, session } = importDocx(
      makeDocx(
        P(
          run("before") +
            control(run("held") + control("", { id: 2 }), { id: 1 })
        )
      )
    );

    const out = documentXmlOf(withEditedBlock(doc, 0, "Edited"), session);

    expect(out).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>' +
        '<w:r><w:t xml:space="preserve">held</w:t></w:r>' +
        '<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>' +
        "</w:sdtContent></w:sdt></w:sdtContent></w:sdt>"
    );
  });

  /**
   * A control nobody rewrote goes out as the bytes it arrived as, so this is the copy: the second
   * one carries an id of its own (§17.5.2.18) and is written from its attrs (`docx/identities`).
   */
  it("gives a second copy of one an id of its own", () => {
    const { doc, session } = importDocx(
      makeDocx(P(run("before") + control("", { id: 9 })))
    );
    const paragraph = doc.child(0);
    const twice = paragraph.copy(
      paragraph.content.append(
        paragraph.content.cut(paragraph.child(0).nodeSize)
      )
    );

    const out = documentXmlOf(withBlocks(doc, [twice]), session);
    const ids = [
      ...out.matchAll(/<w:sdt><w:sdtPr><w:id w:val="(\d+)"\/>/g),
    ].map((match) => match[1]);

    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe("9");
    expect(ids[1]).not.toBe("9");
  });
});
