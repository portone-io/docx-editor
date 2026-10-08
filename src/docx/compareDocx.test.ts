// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { commentedDocx, commentedRun, run } from "../__testing__/comments";
import {
  fixtureNames,
  importErrorCode,
  LETTER_LANDSCAPE_SECT_PR,
  LETTER_SECT_PR,
  makeDocx,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
} from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import { addComment, updateComment } from "../editor/commands/commentCommands";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { onlyCommentsChangedBy } from "./commentOnlyChange";
import { BLOCK_KINDS, compareDocx } from "./compareDocx";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

function paragraph(text: string, rPr = ""): string {
  return `<w:p><w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function body(...blocks: readonly string[]): Uint8Array {
  return makeDocx(blocks.join("") + LETTER_SECT_PR);
}

/** A package declaring its main part, which the writer needs before it can add a comments part */
function withContentTypes(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  parts["[Content_Types].xml"] = new TextEncoder().encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function table(rows: readonly (readonly string[])[]): string {
  const cell = (text: string) => `<w:tc>${paragraph(text)}</w:tc>`;
  return `<w:tbl>${rows.map((cells) => `<w:tr>${cells.map(cell).join("")}</w:tr>`).join("")}</w:tbl>`;
}

describe("compareDocx", () => {
  const lanes = [
    ...fixtureNames.map((name) => ({ name, read: readFixture })),
    ...producerFixtureNames.map((name) => ({
      name,
      read: readProducerFixture,
    })),
  ];

  it.each(lanes)(
    "reports nothing for $name against its own export",
    ({ name, read }) => {
      const bytes = read(name);
      const { doc, session } = importDocx(bytes);
      expect(compareDocx(bytes, exportDocx(doc, session))).toEqual({
        blocks: [],
        comments: [],
        parts: [],
      });
    }
  );

  it.each(lanes)(
    "reports one changed paragraph for $name after one edit",
    ({ name, read }) => {
      const bytes = read(name);
      const { doc, session } = importDocx(bytes);
      let at = -1;
      doc.forEach((node, pos) => {
        if (at === -1 && node.type.name === "paragraph") at = pos;
      });
      expect(at).toBeGreaterThanOrEqual(0);
      const edited = doc.replace(
        at + 1,
        at + 1,
        new Slice(Fragment.from(docxSchema.text("Q!")), 0, 0)
      );
      expect(compareDocx(bytes, exportDocx(edited, session))).toEqual({
        blocks: [
          {
            kind: "changed",
            block: "paragraph",
            original: { pos: at, text: expect.any(String) },
            revised: { pos: at, text: expect.stringMatching(/^Q!/) },
            change: "text",
            edits: expect.arrayContaining([{ kind: "added", text: "Q!" }]),
          },
        ],
        comments: [],
        parts: [],
      });
    }
  );

  it("names a block kind for every block node the schema holds", () => {
    const blocks = Object.values(docxSchema.nodes)
      .filter((type) => type.isInGroup("block"))
      .map((type) => type.name);
    expect(blocks.length).toBeGreaterThan(0);
    expect(Object.keys(BLOCK_KINDS).sort()).toEqual(blocks.sort());
  });

  it("reports a paragraph whose text was edited, where it stands on each side", () => {
    const original = body(
      paragraph("One"),
      paragraph("Two"),
      paragraph("Three")
    );
    const revised = body(
      paragraph("One"),
      paragraph("Two!"),
      paragraph("Three")
    );
    expect(compareDocx(original, revised).blocks).toEqual([
      {
        kind: "changed",
        block: "paragraph",
        original: { pos: 5, text: "Two" },
        revised: { pos: 5, text: "Two!" },
        change: "text",
        edits: [
          { kind: "kept", text: "Two" },
          { kind: "added", text: "!" },
        ],
      },
    ]);
  });

  it("reports an added paragraph at its position in the revised file", () => {
    const original = body(paragraph("One"), paragraph("Three"));
    const revised = body(
      paragraph("One"),
      paragraph("Two"),
      paragraph("Three")
    );
    expect(compareDocx(original, revised).blocks).toEqual([
      { kind: "added", block: "paragraph", pos: 5, text: "Two" },
    ]);
  });

  it("reports a removed paragraph at its position in the original file", () => {
    const original = body(
      paragraph("One"),
      paragraph("Two"),
      paragraph("Three")
    );
    const revised = body(paragraph("One"), paragraph("Three"));
    expect(compareDocx(original, revised).blocks).toEqual([
      { kind: "removed", block: "paragraph", pos: 5, text: "Two" },
    ]);
  });

  it("reports a paragraph that reads the same but gained bold as a formatting change", () => {
    const original = body(paragraph("One"), paragraph("Two"));
    const revised = body(
      paragraph("One"),
      paragraph("Two", "<w:rPr><w:b/></w:rPr>")
    );
    expect(compareDocx(original, revised).blocks).toEqual([
      {
        kind: "changed",
        block: "paragraph",
        original: { pos: 5, text: "Two" },
        revised: { pos: 5, text: "Two" },
        change: "formatting",
      },
    ]);
  });

  it("lists the rows of a changed table: a cell edited and a row added", () => {
    const original = body(
      table([
        ["a", "b"],
        ["c", "d"],
        ["g", "h"],
      ])
    );
    const revised = body(
      table([
        ["a", "b"],
        ["c", "D"],
        ["g", "h"],
        ["e", "f"],
      ])
    );
    const [change] = compareDocx(original, revised).blocks;
    expect(change).toMatchObject({
      kind: "changed",
      block: "table",
      original: { pos: 0, text: "a\nb\nc\nd\ng\nh" },
      revised: { pos: 0, text: "a\nb\nc\nD\ng\nh\ne\nf" },
      change: "text",
      edits: [
        { kind: "kept", text: "a\nb\nc\n" },
        { kind: "removed", text: "d" },
        { kind: "added", text: "D" },
        { kind: "kept", text: "\ng\nh" },
        { kind: "added", text: "\ne\nf" },
      ],
      rows: [
        {
          kind: "changed",
          original: { index: 1, cells: ["c", "d"] },
          revised: { index: 1, cells: ["c", "D"] },
        },
        { kind: "added", index: 3, cells: ["e", "f"] },
      ],
    });
  });

  it("pairs an edited table with its original past a paragraph removed before it", () => {
    const original = body(
      paragraph("The table below"),
      table([
        ["a", "b"],
        ["c", "d"],
      ])
    );
    const revised = body(
      table([
        ["Qa", "b"],
        ["c", "d"],
      ])
    );
    expect(compareDocx(original, revised).blocks).toEqual([
      {
        kind: "removed",
        block: "paragraph",
        pos: 0,
        text: "The table below",
      },
      {
        kind: "changed",
        block: "table",
        original: { pos: 17, text: "a\nb\nc\nd" },
        revised: { pos: 0, text: "Qa\nb\nc\nd" },
        change: "text",
        edits: [
          { kind: "added", text: "Q" },
          { kind: "kept", text: "a\nb\nc\nd" },
        ],
        rows: [
          {
            kind: "changed",
            original: { index: 0, cells: ["a", "b"] },
            revised: { index: 0, cells: ["Qa", "b"] },
          },
        ],
      },
    ]);
  });

  it("reads a table's text as its paragraphs one per line, a cell of two paragraphs taking two", () => {
    const twoParagraphs = `<w:tbl><w:tr><w:tc>${paragraph("a")}${paragraph("b")}</w:tc><w:tc>${paragraph("c")}</w:tc></w:tr></w:tbl>`;
    const [change] = compareDocx(
      body(twoParagraphs),
      body(twoParagraphs.replace(">c<", ">C<"))
    ).blocks;
    expect(change).toMatchObject({
      block: "table",
      original: { text: "a\nb\nc" },
      rows: [
        {
          kind: "changed",
          original: { index: 0, cells: ["a\nb", "c"] },
          revised: { index: 0, cells: ["a\nb", "C"] },
        },
      ],
    });
  });

  it("reports a table whose cells only gained bold as a formatting change with no rows", () => {
    const bold = "<w:rPr><w:b/></w:rPr>";
    const plain = `<w:tbl><w:tr><w:tc>${paragraph("a")}</w:tc></w:tr></w:tbl>`;
    expect(
      compareDocx(body(plain), body(plain.replace("<w:r>", `<w:r>${bold}`)))
        .blocks
    ).toEqual([
      {
        kind: "changed",
        block: "table",
        original: { pos: 0, text: "a" },
        revised: { pos: 0, text: "a" },
        change: "formatting",
        rows: [],
      },
    ]);
  });

  it("pairs two neighbouring edited paragraphs in order", () => {
    const original = body(
      paragraph("One"),
      paragraph("Two"),
      paragraph("Three")
    );
    const revised = body(
      paragraph("Uno"),
      paragraph("Dos"),
      paragraph("Three")
    );
    expect(compareDocx(original, revised).blocks).toMatchObject([
      {
        kind: "changed",
        change: "text",
        original: { text: "One" },
        revised: { text: "Uno" },
      },
      {
        kind: "changed",
        change: "text",
        original: { text: "Two" },
        revised: { text: "Dos" },
      },
    ]);
  });

  it("names a changed styles part and no block for it", () => {
    const text = paragraph("One") + LETTER_SECT_PR;
    const comparison = compareDocx(
      makeDocx(text, '<w:sz w:val="20"/>'),
      makeDocx(text, '<w:sz w:val="28"/>')
    );
    expect(comparison.blocks).toEqual([]);
    expect(comparison.parts).toEqual([
      { part: "word/styles.xml", kind: "changed" },
    ]);
  });

  it("names the main part when the section closing the body changed", () => {
    const comparison = compareDocx(
      makeDocx(paragraph("One") + LETTER_SECT_PR),
      makeDocx(paragraph("One") + LETTER_LANDSCAPE_SECT_PR)
    );
    expect(comparison.blocks).toEqual([]);
    expect(comparison.parts).toEqual([
      { part: "word/document.xml", kind: "changed" },
    ]);
  });

  it("turns down bytes that are not a docx the way opening them does", () => {
    expect(
      importErrorCode(() =>
        compareDocx(new Uint8Array([1, 2, 3]), body(paragraph("One")))
      )
    ).toBe(importErrorCode(() => importDocx(new Uint8Array([1, 2, 3]))));
  });
});

describe("compareDocx over comments", () => {
  const author = { id: "me", name: "Me" };

  function edited(
    bytes: Uint8Array,
    edit: (state: EditorState) => EditorState
  ): Uint8Array {
    const { doc, session } = importDocx(bytes);
    return exportDocx(edit(createEditorState(doc, { author })).doc, session);
  }

  function commented(state: EditorState, text: string): EditorState {
    const { from, to } = rangeOfText(state.doc, "beta");
    let next = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to))
    );
    addComment({ text, author: "Me", authorId: "me" })(
      next,
      (tr) => (next = next.apply(tr))
    );
    return next;
  }

  const original = withContentTypes(
    body(paragraph("Alpha beta"), paragraph("Gamma"))
  );

  it("reports a comment added through the editor, and no block for it", () => {
    const submitted = edited(original, (state) => commented(state, "note"));
    const comparison = compareDocx(original, submitted);
    expect(comparison.blocks).toEqual([]);
    // The first comment of a file gains a relationship and a content type for its parts
    expect(comparison.parts).toEqual([]);
    expect(comparison.comments).toEqual([
      {
        kind: "added",
        id: expect.any(String),
        author: "Me",
        authorId: "me",
        text: "note",
      },
    ]);
  });

  describe("in a paragraph holding a run with no characters", () => {
    const slot =
      '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>';
    const withSlot = withContentTypes(
      body(
        `<w:p><w:r><w:t xml:space="preserve">Alpha beta </w:t></w:r>${slot}</w:p>`,
        paragraph("Gamma")
      )
    );

    function commentedOver(
      state: EditorState,
      range: (state: EditorState) => { from: number; to: number }
    ): EditorState {
      const { from, to } = range(state);
      let next = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, from, to))
      );
      addComment({ text: "note", author: "Me", authorId: "me" })(
        next,
        (tr) => (next = next.apply(tr))
      );
      return next;
    }

    it.each([
      [
        "a word beside it",
        (state: EditorState) => rangeOfText(state.doc, "beta"),
      ],
      [
        "the run itself",
        (state: EditorState) => {
          let at = -1;
          state.doc.descendants((node, pos) => {
            if (node.type.name === "emptyRun") at = pos;
          });
          return { from: at, to: at + 1 };
        },
      ],
    ])(
      "reports a comment over %s and no block, and passes as comments alone",
      (_, range) => {
        const submitted = edited(withSlot, (state) =>
          commentedOver(state, range)
        );
        const comparison = compareDocx(withSlot, submitted);
        expect(comparison.blocks).toEqual([]);
        expect(comparison.comments).toHaveLength(1);
        expect(onlyCommentsChangedBy(withSlot, submitted, "me")).toEqual({
          ok: true,
        });
      }
    );
  });

  it("reports a comment removed, and no block", () => {
    const submitted = edited(original, (state) => commented(state, "note"));
    const comparison = compareDocx(submitted, original);
    expect(comparison.blocks).toEqual([]);
    expect(comparison.parts).toEqual([]);
    expect(comparison.comments).toMatchObject([
      { kind: "removed", author: "Me", text: "note" },
    ]);
  });

  it("reports a comment whose text was rewritten", () => {
    const first = edited(original, (state) => commented(state, "note"));
    const [added] = compareDocx(original, first).comments;
    const second = edited(first, (state) => {
      let next = state;
      updateComment(added.id, "revised")(next, (tr) => (next = next.apply(tr)));
      return next;
    });
    expect(compareDocx(first, second)).toEqual({
      blocks: [],
      comments: [
        {
          kind: "changed",
          id: added.id,
          author: "Me",
          authorId: "me",
          original: "note",
          revised: "revised",
        },
      ],
      parts: [],
    });
  });

  it("still names the relationships and content types beside a comment when they gained something else", () => {
    const submitted = edited(original, (state) => commented(state, "note"));
    const parts = unzipSync(submitted);
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    parts["[Content_Types].xml"] = encoder.encode(
      decoder
        .decode(parts["[Content_Types].xml"])
        .replace(
          "</Types>",
          '<Default Extension="png" ContentType="image/png"/></Types>'
        )
    );
    parts["word/_rels/document.xml.rels"] = encoder.encode(
      decoder
        .decode(parts["word/_rels/document.xml.rels"])
        .replace(
          "</Relationships>",
          '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>'
        )
    );
    expect(compareDocx(original, zipSync(parts)).parts).toEqual([
      { part: "[Content_Types].xml", kind: "changed" },
      { part: "word/_rels/document.xml.rels", kind: "added" },
    ]);
  });

  it("reports a comment under a reused id by another author as removed and added", () => {
    const original = commentedDocx(
      `<w:p>${commentedRun("0", "Alpha")}${run(" beta")}</w:p>${LETTER_SECT_PR}`,
      [{ id: "0", text: "first", author: "Ann" }]
    );
    const revised = commentedDocx(
      `<w:p>${commentedRun("0", "Alpha")}${commentedRun("1", " beta")}</w:p>${LETTER_SECT_PR}`,
      [
        { id: "0", text: "new", author: "Bob" },
        { id: "1", text: "first", author: "Ann" },
      ]
    );
    expect(compareDocx(original, revised).comments).toEqual([
      {
        kind: "removed",
        id: "0",
        author: "Ann",
        authorId: null,
        text: "first",
      },
      { kind: "added", id: "0", author: "Bob", authorId: null, text: "new" },
      { kind: "added", id: "1", author: "Ann", authorId: null, text: "first" },
    ]);
  });

  describe("names the content types beside a comment", () => {
    const COMMENTS_TYPE =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

    function declared(
      bytes: Uint8Array,
      rewrite: (types: string) => string
    ): Uint8Array {
      const parts = unzipSync(bytes);
      parts["[Content_Types].xml"] = new TextEncoder().encode(
        rewrite(new TextDecoder().decode(parts["[Content_Types].xml"]))
      );
      return zipSync(parts);
    }

    const commentedFile = () =>
      commentedDocx(
        `<w:p>${commentedRun("1", "Alpha")}</w:p>${LETTER_SECT_PR}`,
        [{ id: "1", text: "first", author: "Ann" }]
      );

    it("when the comments part lost its override but is still there", () => {
      const bytes = commentedFile();
      const undeclared = declared(bytes, (types) =>
        types.replace(/<Override PartName="\/word\/comments\.xml"[^>]*\/>/, "")
      );
      expect(compareDocx(bytes, undeclared).parts).toEqual([
        { part: "[Content_Types].xml", kind: "changed" },
      ]);
    });

    it("when another part is declared as comments", () => {
      const bytes = commentedFile();
      const retyped = declared(bytes, (types) =>
        types.replace(
          "</Types>",
          `<Override PartName="/word/styles.xml" ContentType="${COMMENTS_TYPE}"/></Types>`
        )
      );
      expect(compareDocx(bytes, retyped).parts).toEqual([
        { part: "[Content_Types].xml", kind: "changed" },
      ]);
    });

    it("when an extension is declared as comments", () => {
      const bytes = commentedFile();
      const byExtension = declared(bytes, (types) =>
        types.replace(
          "</Types>",
          `<Default Extension="bin" ContentType="${COMMENTS_TYPE}"/></Types>`
        )
      );
      expect(compareDocx(bytes, byExtension).parts).toEqual([
        { part: "[Content_Types].xml", kind: "changed" },
      ]);
    });
  });

  it("matches comments by id across two files holding different ones", () => {
    const original = commentedDocx(
      `<w:p>${commentedRun("1", "Alpha")}${run(" beta")}</w:p>${LETTER_SECT_PR}`,
      [{ id: "1", text: "first", author: "Ann" }]
    );
    const revised = commentedDocx(
      `<w:p>${run("Alpha ")}${commentedRun("2", "beta")}</w:p>${LETTER_SECT_PR}`,
      [{ id: "2", text: "second" }]
    );
    const comparison = compareDocx(original, revised);
    expect(comparison.comments).toEqual([
      {
        kind: "removed",
        id: "1",
        author: "Ann",
        authorId: null,
        text: "first",
      },
      { kind: "added", id: "2", author: null, authorId: null, text: "second" },
    ]);
    expect(comparison.parts).toEqual([]);
  });
});
