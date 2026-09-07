// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeDocx,
  makeNumberedDocx,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { rangeOfText, select } from "../__testing__/editing";
import {
  addComment,
  setCommentResolved,
} from "../editor/commands/commentCommands";
import { toggleNumberedList } from "../editor/commands/listCommands";
import {
  createEditorState,
  editorStateForSession,
} from "../editor/createEditor";
import type { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import { commentReferencesIn } from "./comments";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { exportProblems } from "./invariants";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const paragraph = (text: string) => `<w:p>${run(text)}</w:p>`;

const cell = (text: string, props = "") =>
  `<w:tc>${props === "" ? "" : `<w:tcPr>${props}</w:tcPr>`}<w:p>${run(text)}</w:p></w:tc>`;

/** A two by two table whose first column is one vertical merge */
const MERGED_TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cell("Party", '<w:vMerge w:val="restart"/>')}${cell("Name")}</w:tr>` +
  `<w:tr>${cell("", "<w:vMerge/>")}${cell("Address")}</w:tr>` +
  "</w:tbl>";

const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** A package whose Comments part has a root the writer cannot rewrite around */
function rootlessCommentsDocx(body: string): Uint8Array {
  const encoder = new TextEncoder();
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(
    `<w:notes xmlns:w="${W_NS}"><w:comment w:id="1"/></w:notes>`
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function apply(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(
    command(state, (tr) => {
      next = state.apply(tr);
    })
  ).toBe(true);
  return next;
}

/** The document with the first paragraph turned into a list, which is a new list wherever the document defines none */
function withNewList(doc: PMNode): PMNode {
  const state = select(createEditorState(doc), 1);
  return apply(state, toggleNumberedList).doc;
}

/** The document with its first paragraph carrying an image that was not in the file */
function withInsertedImage(doc: PMNode): PMNode {
  const first = doc.child(0);
  const image = docxSchema.nodes.image.create({
    src: TINY_PNG_DATA_URL,
    extent: { cx: 9525, cy: 9525 },
  });
  return doc.copy(
    Fragment.from([
      first.copy(Fragment.from([...first.children, image])),
      ...doc.children.slice(1),
    ])
  );
}

/** The state with a comment written over the first word */
function withComment(state: EditorState): EditorState {
  const { from, to } = rangeOfText(state.doc, "Body");
  return apply(
    select(state, from, to),
    addComment({ text: "Check this", author: "Ada" })
  );
}

describe("bookmark pairs", () => {
  it.each([
    '<w:customXml><!-- <w:bookmarkStart w:id="7"/> --><w:p/></w:customXml>',
    '<w:customXml><![CDATA[<w:bookmarkStart w:id="7"/>]]><w:p/></w:customXml>',
    '<w:customXml><x:bookmarkStart xmlns:x="urn:foreign" x:id="7"/><w:p/></w:customXml>',
    '<w:bookmarkStart w:name="a>b" w:id="7"/><w:p/><w:bookmarkEnd w:id="7"/>',
    '<w:bookmarkStart w:id="&#55;" w:name="Name"/><w:p/><w:bookmarkEnd w:id="7"/>',
  ])("agrees with the writer about actual bookmark elements: %s", (body) => {
    const opened = importDocx(makeDocx(body));
    expect(exportProblems(opened.doc, opened.session)).toEqual([]);
    expect(() => exportDocx(opened.doc, opened.session)).not.toThrow();
  });

  it("an unpaired bookmark start is a malformed-xml problem at its position", () => {
    const opened = importDocx(
      makeDocx(
        paragraph("Before") +
          '<w:bookmarkStart w:id="8" w:name="Appendix"/>' +
          paragraph("Inside") +
          '<w:bookmarkEnd w:id="8"/>'
      )
    );
    const unpaired = opened.doc.copy(
      Fragment.from(opened.doc.children.slice(0, -1))
    );

    expect(exportProblems(unpaired, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "bookmark 8 has no end marker",
        pos: opened.doc.child(0).nodeSize,
      },
    ]);
  });

  it("a second start of the same bookmark is reported where the second one stands", () => {
    const opened = importDocx(
      makeDocx(
        '<w:bookmarkStart w:id="8" w:name="Appendix"/>' +
          paragraph("Inside") +
          '<w:bookmarkEnd w:id="8"/>'
      )
    );
    const [start, inside, end] = opened.doc.children;
    if (!start || !inside || !end) throw new Error("three blocks expected");
    const doubled = opened.doc.copy(Fragment.from([start, inside, start, end]));

    expect(exportProblems(doubled, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "bookmark 8 has more than one start marker",
        pos: start.nodeSize + inside.nodeSize,
      },
    ]);
  });

  it("a marker inside a paragraph that ends what nothing started is reported where the marker stands", () => {
    const opened = importDocx(
      makeDocx(`<w:p>${run("Inside")}<w:bookmarkEnd w:id="9"/></w:p>`)
    );

    expect(exportProblems(opened.doc, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "bookmark 9 ends without an earlier start marker",
        pos: 1 + "Inside".length,
      },
    ]);
  });
});

describe("table grids", () => {
  it("a vertical merge past the last row is an invalid-table problem", () => {
    const opened = importDocx(makeDocx(paragraph("Before") + MERGED_TABLE));
    const table = opened.doc.child(1);
    const rows = table.children.map((row, index) =>
      index === 0
        ? row.copy(
            Fragment.from(
              row.children.map((entry, column) =>
                column === 0
                  ? entry.type.create(
                      { ...entry.attrs, rowspan: 9 },
                      entry.content
                    )
                  : entry
              )
            )
          )
        : row
    );
    const overlong = opened.doc.copy(
      Fragment.from([opened.doc.child(0), table.copy(Fragment.from(rows))])
    );

    expect(exportProblems(overlong, opened.session)).toEqual([
      {
        code: "invalid-table",
        message: "a vertical merge in the table reaches past the last row",
        pos: opened.doc.child(0).nodeSize,
      },
    ]);
  });
});

describe("preserved originals", () => {
  it("a preserved element that lost its XML is a lost-original problem", () => {
    const opened = importDocx(makeDocx(paragraph("Body")));
    const stripped = opened.doc.copy(
      Fragment.from([
        opened.doc.child(0),
        docxSchema.nodes.rawBlock.create({ name: "w:tbl" }),
      ])
    );

    expect(exportProblems(stripped, opened.session)).toEqual([
      {
        code: "lost-original",
        message: "a preserved element has lost its original XML",
        pos: opened.doc.child(0).nodeSize,
      },
    ]);
  });

  it("a placeholder from another session is a lost-original problem", () => {
    const alpha = importDocx(
      makeDocx(
        paragraph("alpha") +
          '<w:customXml w:uri="urn:alpha" w:element="from-alpha"/>'
      )
    );
    const beta = importDocx(makeDocx(paragraph("beta")));
    const moved = beta.doc.copy(
      Fragment.from([beta.doc.child(0), alpha.doc.child(1)])
    );

    expect(exportProblems(moved, beta.session)).toEqual([
      {
        code: "lost-original",
        message: `a preserved block comes from another document (${alpha.session.sessionId})`,
        pos: beta.doc.child(0).nodeSize,
      },
    ]);
  });
});

describe("the numbering part", () => {
  it("a new list in a document without numbering.xml is a missing-numbering-part problem", () => {
    const opened = importDocx(makeDocx(paragraph("Body")));

    expect(exportProblems(withNewList(opened.doc), opened.session)).toEqual([
      {
        code: "missing-numbering-part",
        message:
          "cannot add a new list to a document that has no numbering.xml",
      },
    ]);
  });

  it("a new list in a document with numbering.xml is no problem", () => {
    const opened = importDocx(makeNumberedDocx(paragraph("Body")));
    const listed = apply(
      select(editorStateForSession(opened), 1),
      toggleNumberedList
    );

    expect(exportProblems(listed.doc, opened.session)).toEqual([]);
  });
});

describe("content types", () => {
  it("an inserted image without a content types part is a missing-content-types problem", () => {
    const opened = importDocx(makeDocx(paragraph("Body")));

    expect(
      exportProblems(withInsertedImage(opened.doc), opened.session)
    ).toEqual([
      {
        code: "missing-content-types",
        message:
          "cannot add an image to a package that has no [Content_Types].xml",
      },
    ]);
  });

  it("reports missing content types when an existing comment gains its first thread part", () => {
    const parts = unzipSync(rootlessCommentsDocx(paragraph("Body")));
    parts["word/comments.xml"] = new TextEncoder().encode(
      `<w:comments xmlns:w="${W_NS}"/>`
    );
    const opened = importDocx(zipSync(parts));
    const commented = withComment(editorStateForSession(opened));
    const saved = unzipSync(exportDocx(commented.doc, opened.session));
    delete saved["[Content_Types].xml"];
    const reopened = importDocx(zipSync(saved));
    const id = commentReferencesIn(reopened.doc).keys().next().value;
    expect(id).toBeDefined();
    const resolved = apply(
      editorStateForSession(reopened),
      setCommentResolved(id ?? "", true)
    );
    expect(exportProblems(resolved.doc, reopened.session)).toEqual([
      {
        code: "missing-content-types",
        message:
          "cannot add a part to a package that has no [Content_Types].xml",
      },
    ]);
  });

  it("a first comment in a package without a content types part is a missing-content-types problem", () => {
    const opened = importDocx(makeDocx(paragraph("Body")));
    const commented = withComment(editorStateForSession(opened));

    expect(exportProblems(commented.doc, opened.session)).toEqual([
      {
        code: "missing-content-types",
        message:
          "cannot add a part to a package that has no [Content_Types].xml",
      },
    ]);
  });
});

describe("comment part roots", () => {
  it("a rootless comments part is reported only once a comment changes", () => {
    const opened = importDocx(rootlessCommentsDocx(paragraph("Body")));
    const untouched = editorStateForSession(opened);

    expect(exportProblems(untouched.doc, opened.session)).toEqual([]);
    expect(exportProblems(withComment(untouched).doc, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "the Comments part has no comments root element",
      },
    ]);
  });
});

describe("a document nobody touched", () => {
  it.each(fixtureNames)("%s: opened and untouched has no problem", (name) => {
    const { doc, session } = importDocx(readFixture(name));
    expect(exportProblems(doc, session)).toEqual([]);
  });

  it.each(producerFixtureNames)(
    "%s: opened and untouched has no problem",
    (name) => {
      const { doc, session } = importDocx(readProducerFixture(name));
      expect(exportProblems(doc, session)).toEqual([]);
    }
  );
});

describe("exportDocx", () => {
  it("throws the first problem with its code", () => {
    const alpha = importDocx(
      makeDocx(
        paragraph("alpha") +
          '<w:customXml w:uri="urn:alpha" w:element="from-alpha"/>'
      )
    );
    const opened = importDocx(makeDocx(paragraph("Body")));
    // The writer would meet the image's missing content types before it reached the block
    // from elsewhere, so the list's order is what has to be thrown, not the writer's
    const troubled = withInsertedImage(
      opened.doc.copy(Fragment.from([opened.doc.child(0), alpha.doc.child(1)]))
    );
    const problems = exportProblems(troubled, opened.session);
    expect(problems.map((problem) => problem.code)).toEqual([
      "lost-original",
      "missing-content-types",
    ]);

    expect(() => exportDocx(troubled, opened.session)).toThrowError(
      expect.objectContaining<Partial<DocxExportError>>({
        code: problems[0]?.code,
        message: problems[0]?.message,
      })
    );
  });
});
