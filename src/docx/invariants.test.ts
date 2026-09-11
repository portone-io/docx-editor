// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeDeclaredDocx,
  makeDocx,
  makeHeadersFootersDocx,
  makeNotesDocx,
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
import {
  STORIES_ATTR,
  type StoryKey,
  storiesOf,
  storyKey,
  storyNodeOf,
} from "../schema/stories";
import { withEditedFirst } from "./__testing__/blockEdits";
import { commentReferencesIn } from "./comments";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { exportProblems } from "./invariants";
import { storyFromText } from "./story";

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
    [
      "a cell paragraph and its following marker",
      1,
      '<w:tr><w:tc><w:p><w:bookmarkStart w:id="4" w:name="Range"/>' +
        run("a") +
        '</w:p></w:tc><w:bookmarkEnd w:id="4"/></w:tr>',
    ],
    [
      "a cell paragraph and a following row boundary",
      1,
      '<w:tr><w:tc><w:p><w:bookmarkStart w:id="4" w:name="Range"/>' +
        run("a") +
        '</w:p></w:tc></w:tr><w:bookmarkEnd w:id="4"/>',
    ],
    [
      "two cells whose respective markers cross the cell boundary",
      2,
      "<w:tr><w:tc>" +
        paragraph("a") +
        '</w:tc><w:bookmarkStart w:id="4" w:name="Range"/>' +
        "<w:tc><w:p>" +
        run("b") +
        '<w:bookmarkEnd w:id="4"/></w:p></w:tc></w:tr>',
    ],
  ])("reads bookmark order across %s", (_name, columns, rows) => {
    const opened = importDocx(
      makeDocx(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/>' +
          (columns === 2 ? '<w:gridCol w:w="1000"/>' : "") +
          "</w:tblGrid>" +
          rows +
          "</w:tbl>"
      )
    );
    expect(opened.doc.firstChild?.type.name).toBe("table");
    for (const doc of [
      opened.doc,
      withEditedFirst(opened.doc, "table", "Edited"),
    ]) {
      expect(exportProblems(doc, opened.session)).toEqual([]);
      expect(() => exportDocx(doc, opened.session)).not.toThrow();
    }
  });

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

  it("does not let a later cell attribute satisfy an earlier inline end", () => {
    const opened = importDocx(
      makeDocx(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
          '<w:tr><w:tc><w:p><w:bookmarkEnd w:id="4"/>' +
          run("a") +
          '</w:p></w:tc><w:bookmarkStart w:id="4" w:name="Range"/></w:tr></w:tbl>'
      )
    );
    expect(
      exportProblems(opened.doc, opened.session).map(({ message }) => message)
    ).toEqual([
      "bookmark 4 ends without an earlier start marker",
      "bookmark 4 has no end marker",
    ]);
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

    // The doubled marker is the one preserved block standing in two places as well, which the
    // identity pass reports at the same place after the pairing
    expect(exportProblems(doubled, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "bookmark 8 has more than one start marker",
        pos: start.nodeSize + inside.nodeSize,
      },
      {
        code: "unsupported-content",
        message: "a preserved block stands in two places (rawBlock)",
        pos: start.nodeSize + inside.nodeSize,
      },
    ]);
  });

  it("a row-level marker that ends what nothing started is reported at its carrier", () => {
    const opened = importDocx(
      makeDocx(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
          `<w:tr><w:tc>${paragraph("a")}</w:tc>` +
          '<w:bookmarkEnd w:id="4"/></w:tr></w:tbl>'
      )
    );

    expect(exportProblems(opened.doc, opened.session)).toEqual([
      {
        code: "malformed-xml",
        message: "bookmark 4 ends without an earlier start marker",
        pos: 2,
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
        message: "a preserved block has lost its original XML",
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

describe("unique identities", () => {
  it("a preserved block standing twice is an unsupported-content problem where the second one stands", () => {
    const opened = importDocx(
      makeDocx(
        paragraph("Body") +
          '<w:customXml w:uri="urn:placeholder" w:element="kept"/>'
      )
    );
    const placeholder = opened.doc.child(1);
    const twice = opened.doc.copy(
      Fragment.from([opened.doc.child(0), placeholder, placeholder])
    );

    const problems = exportProblems(twice, opened.session);
    expect(problems).toEqual([
      {
        code: "unsupported-content",
        message: "a preserved block stands in two places (rawBlock)",
        pos: opened.doc.child(0).nodeSize + placeholder.nodeSize,
      },
    ]);
    expect(() => exportDocx(twice, opened.session)).toThrowError(
      expect.objectContaining<Partial<DocxExportError>>({
        code: problems[0]?.code,
        message: problems[0]?.message,
      })
    );
  });
});

describe("the numbering part", () => {
  it("a new list in a document the export can write the part for is no problem", () => {
    const opened = importDocx(makeDeclaredDocx(paragraph("Body")));

    expect(exportProblems(withNewList(opened.doc), opened.session)).toEqual([]);
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

  it("a new list in a package without a content types part is a missing-content-types problem", () => {
    const opened = importDocx(makeDocx(paragraph("Body")));

    expect(exportProblems(withNewList(opened.doc), opened.session)).toEqual([
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
        message: "the comments part has no comments root element",
      },
    ]);
  });
});

describe("unique identities in a header", () => {
  it("a preserved block standing twice in an edited header is an unsupported-content problem with no position", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header1.xml"] = new TextEncoder().encode(
      `<w:hdr xmlns:w="${W_NS}">${paragraph("Head")}` +
        '<w:customXml w:uri="urn:placeholder" w:element="kept"/></w:hdr>'
    );
    const opened = importDocx(zipSync(parts));
    const key = storyKey("header", "word/header1.xml");
    const story = storyNodeOf(opened.doc, key);
    if (!story) throw new Error("the header was not read");
    const placeholder = story.child(1);
    const twice = story.copy(
      Fragment.from([story.child(0), placeholder, placeholder])
    );
    const edited = opened.doc.type.create(
      {
        ...opened.doc.attrs,
        [STORIES_ATTR]: { ...storiesOf(opened.doc), [key]: twice.toJSON() },
      },
      opened.doc.content
    );

    const problems = exportProblems(edited, opened.session);
    expect(problems).toEqual([
      {
        code: "unsupported-content",
        message: "a preserved block stands in two places (rawBlock)",
      },
    ]);
    expect(() => exportDocx(edited, opened.session)).toThrowError(
      expect.objectContaining<Partial<DocxExportError>>({
        code: problems[0]?.code,
        message: problems[0]?.message,
      })
    );
  });
});

/** The document with this story standing under its key, the way an edit writes one */
function withStory(doc: PMNode, key: StoryKey, story: PMNode): PMNode {
  return doc.type.create(
    {
      ...doc.attrs,
      [STORIES_ATTR]: { ...storiesOf(doc), [key]: story.toJSON() },
    },
    doc.content
  );
}

/** The problems the document reports, and that the export throws the first of them */
function refusedWith(doc: PMNode, opened: ReturnType<typeof importDocx>) {
  const problems = exportProblems(doc, opened.session);
  expect(() => exportDocx(doc, opened.session)).toThrowError(
    expect.objectContaining<Partial<DocxExportError>>({
      code: problems[0]?.code,
      message: problems[0]?.message,
    })
  );
  return problems;
}

describe("stories no part writer writes", () => {
  it("refuses an export that changed an endnote story no planner writes", () => {
    const opened = importDocx(makeNotesDocx());
    const edited = withStory(
      opened.doc,
      storyKey("endnote", "3"),
      storyFromText("Rewritten")
    );

    expect(refusedWith(edited, opened)).toEqual([
      {
        code: "unsupported-content",
        message: "the endnote:3 story changed, and no part writer writes it",
      },
    ]);
  });

  it("refuses an export that changed a separator entry, which goes back out as it arrived", () => {
    const opened = importDocx(makeNotesDocx());
    const edited = withStory(
      opened.doc,
      storyKey("footnote", "-1"),
      storyFromText("Not a line")
    );

    expect(refusedWith(edited, opened)).toEqual([
      {
        code: "unsupported-content",
        message: "the footnote:-1 story changed, and no part writer writes it",
      },
    ]);
  });
});

describe("the footnotes part", () => {
  it("refuses to add a footnotes part to a package with no content types", () => {
    const opened = importDocx(
      makeDocx(
        `<w:p>${run("Text")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`
      )
    );
    const edited = withStory(
      opened.doc,
      storyKey("footnote", "1"),
      storyFromText("A new note")
    );

    expect(refusedWith(edited, opened)).toEqual([
      {
        code: "missing-content-types",
        message:
          "cannot add a part to a package that has no [Content_Types].xml",
      },
    ]);
  });

  it("reports a footnotes part with no root once a footnote in it changed", () => {
    const parts = unzipSync(makeNotesDocx());
    parts["word/footnotes.xml"] = new TextEncoder().encode(
      `<w:notes xmlns:w="${W_NS}"><w:footnote w:id="2">${paragraph("Body")}</w:footnote></w:notes>`
    );
    const opened = importDocx(zipSync(parts));

    expect(exportProblems(opened.doc, opened.session)).toEqual([]);
    const edited = withStory(
      opened.doc,
      storyKey("footnote", "2"),
      storyFromText("Rewritten")
    );
    expect(refusedWith(edited, opened)).toEqual([
      {
        code: "malformed-xml",
        message: "the footnotes part has no footnotes root element",
      },
    ]);
  });

  it("a preserved block standing twice in an edited footnote is an unsupported-content problem with no position", () => {
    const parts = unzipSync(makeNotesDocx());
    parts["word/footnotes.xml"] = new TextEncoder().encode(
      `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:id="2">${paragraph("Note")}` +
        '<w:customXml w:uri="urn:placeholder" w:element="kept"/></w:footnote></w:footnotes>'
    );
    const opened = importDocx(zipSync(parts));
    const key = storyKey("footnote", "2");
    const story = storyNodeOf(opened.doc, key);
    if (!story) throw new Error("the footnote was not read");
    const placeholder = story.child(1);
    const edited = withStory(
      opened.doc,
      key,
      story.copy(Fragment.from([story.child(0), placeholder, placeholder]))
    );

    expect(refusedWith(edited, opened)).toEqual([
      {
        code: "unsupported-content",
        message: "a preserved block stands in two places (rawBlock)",
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
