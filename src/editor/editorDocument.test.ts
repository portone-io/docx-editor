// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  LETTER_GEOMETRY,
  LETTER_SECT_PR,
  makeDeclaredDocx,
  makeDocx,
  makeStyledNumberedDocx,
} from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { documentNumbering as sessionNumbering } from "../docx/session";
import { createEditorState, editorStateForSession } from "./createEditor";
import { documentFormatting, documentGeometry } from "./documentStyles";
import {
  documentOf,
  editorDocumentOf,
  NO_DOCUMENT,
  reservedCommentIds,
  reservedCommentParaIds,
} from "./editorDocument";
import {
  canStartNewList,
  documentNumbering,
} from "./plugins/numberingDecorations";

const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_REL_BASE = "http://schemas.microsoft.com/office/2011/relationships";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

const encoder = new TextEncoder();

const BODY =
  '<w:p><w:commentRangeStart w:id="4"/>' +
  '<w:r><w:t xml:space="preserve">Alpha</w:t></w:r>' +
  '<w:commentRangeEnd w:id="4"/>' +
  '<w:r><w:commentReference w:id="4"/></w:r></w:p>';

/** Document defaults, one style to point at, and one style every paragraph wears by default */
const STYLES =
  "<w:docDefaults><w:pPrDefault><w:pPr>" +
  '<w:spacing w:line="360" w:lineRule="auto"/>' +
  "</w:pPr></w:pPrDefault></w:docDefaults>" +
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
  '<w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote">' +
  '<w:name w:val="Quote"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>';

const COMMENTS_XML =
  `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
  '<w:comment w:id="4" w:author="Ada"><w:p w14:paraId="000000A1">' +
  '<w:r><w:t xml:space="preserve">Check this</w:t></w:r></w:p></w:comment>' +
  "</w:comments>";

/** One entry for the comment above, and one whose comment is no longer in the part */
const COMMENTS_EXTENDED_XML =
  `<w15:commentsEx xmlns:w15="${W15_NS}">` +
  '<w15:commentEx w15:paraId="000000A1"/>' +
  '<w15:commentEx w15:paraId="DEADBEEF"/>' +
  "</w15:commentsEx>";

function documentRels(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Target="styles.xml" Type="${REL_BASE}/styles"/>` +
    `<Relationship Id="rId2" Target="numbering.xml" Type="${REL_BASE}/numbering"/>` +
    `<Relationship Id="rId3" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
    `<Relationship Id="rId4" Target="commentsExtended.xml" Type="${MS_REL_BASE}/commentsExtended"/>` +
    `<Relationship Id="rId5" Target="settings.xml" Type="${REL_BASE}/settings"/>` +
    "</Relationships>"
  );
}

/**
 * A document that wrote every document-level value down: a style chain with defaults of its own,
 * list definitions, its own paper, its own tab interval, and a comment part whose identifiers are
 * spent.
 */
function opened() {
  const parts = unzipSync(
    makeStyledNumberedDocx(BODY + LETTER_SECT_PR, STYLES)
  );
  parts["word/_rels/document.xml.rels"] = encoder.encode(documentRels());
  parts["word/comments.xml"] = encoder.encode(COMMENTS_XML);
  parts["word/commentsExtended.xml"] = encoder.encode(COMMENTS_EXTENDED_XML);
  parts["word/settings.xml"] = encoder.encode(
    `<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="960"/>` +
      "<w:compat><w:noTabHangInd/></w:compat></w:settings>"
  );
  return importDocx(zipSync(parts));
}

describe("reading an opened document into editor values", () => {
  it("reads every session field the component used to copy", () => {
    const { session } = opened();
    const document = editorDocumentOf(session);

    expect(document.session).toBe(session);
    expect(document.formatting).toBe(session.formatting);
    expect(document.defaults).toBe(session.defaults);
    expect(document.paragraphStyles).toBe(session.paragraphStyles);
    expect(document.formatting.numbering).toEqual(sessionNumbering(session));
    expect(document.canStartNewList).toBe(true);
    expect(document.geometry).toBe(session.geometry);
    expect(document.defaultTabStopPt).toBe(session.defaultTabStopPt);
    expect(document.reservedCommentIds).toEqual(
      new Set(session.comments.byId.keys())
    );
    expect(document.reservedCommentParaIds).toEqual(
      new Set([
        ...session.comments.ordered.flatMap((comment) =>
          comment.paraId === null ? [] : [comment.paraId]
        ),
        ...session.comments.extendedOrdered.map(
          (extension) => extension.paraId
        ),
      ])
    );

    // Every one of them is a value the document actually wrote down, not a fallback
    expect(document.formatting.defaultParagraphStyleId).toBe("Normal");
    expect(document.formatting.compat).toEqual({ noTabHangInd: true });
    expect(
      document.formatting.paragraphDefaults.lineSpacing
    ).not.toBeUndefined();
    expect(document.formatting.numbering.lists.size).toBeGreaterThan(0);
    expect(document.geometry).toEqual(LETTER_GEOMETRY);
    expect(document.defaultTabStopPt).toBe(48);
    expect([...document.reservedCommentIds]).toEqual(["4"]);
    expect([...document.reservedCommentParaIds]).toEqual([
      "000000A1",
      "DEADBEEF",
    ]);
  });

  it("a state built for a session answers numbering, styles, geometry and reservations from one value", () => {
    const state = editorStateForSession(opened());
    const document = documentOf(state);

    expect(documentNumbering(state)).toBe(document.formatting.numbering);
    expect(canStartNewList(state)).toBe(document.canStartNewList);
    expect(documentFormatting(state)).toBe(document.formatting);
    expect(documentGeometry(state)).toBe(document.geometry);
    expect(reservedCommentIds(state)).toBe(document.reservedCommentIds);
    expect(reservedCommentParaIds(state)).toBe(document.reservedCommentParaIds);

    expect(documentGeometry(state)).toEqual(LETTER_GEOMETRY);
    expect(canStartNewList(state)).toBe(true);
    expect([...reservedCommentIds(state)]).toEqual(["4"]);
  });

  it("a document with no numbering part can still start a list, since the export writes one", () => {
    const declared = importDocx(makeDeclaredDocx(BODY));
    const bare = importDocx(makeDocx(BODY));

    expect(declared.session.numberingPartPath).toBeNull();
    expect(editorDocumentOf(declared.session).canStartNewList).toBe(true);
    // Declaring the part it adds is the one thing a package can leave the export no room for
    expect(editorDocumentOf(bare.session).canStartNewList).toBe(false);
  });

  it("a state built without an opened document reads the empty snapshot", () => {
    const { doc } = opened();

    expect(documentOf(createEditorState(doc))).toBe(NO_DOCUMENT);
  });
});

describe("the snapshot across transactions", () => {
  it("the snapshot is the same object across selection-only transactions", () => {
    const state = editorStateForSession(opened());
    const moved = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1))
    );

    expect(documentOf(moved)).toBe(documentOf(state));
  });

  it("the snapshot is re-derived when the document node's attributes change identity", () => {
    const state = editorStateForSession(opened());
    // Editing the text rebuilds the document node around the same attrs, so nothing is re-derived
    const typed = state.apply(state.tr.insertText("more", 1));
    expect(documentOf(typed)).toBe(documentOf(state));

    const attributed = typed.apply(typed.tr.setDocAttribute("sectPr", null));

    expect(attributed.doc.attrs).not.toBe(typed.doc.attrs);
    expect(documentOf(attributed)).not.toBe(documentOf(typed));
    expect(documentOf(attributed)).toEqual(documentOf(typed));
  });
});
