// @vitest-environment jsdom
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  clipboardSerializer,
  clipboardSpecs,
  clipboardText,
} from "./clipboard";
import { docxSchema } from "./docxSchema";

const TINY_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function paragraph(...content: readonly PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(null, content);
}

function cell(...content: readonly PMNode[]): PMNode {
  return docxSchema.nodes.tableCell.create(null, content);
}

function copiedHtml(...blocks: readonly PMNode[]): string {
  const wrap = document.createElement("div");
  wrap.appendChild(
    clipboardSerializer().serializeFragment(Fragment.from(blocks), {
      document,
    })
  );
  return wrap.innerHTML;
}

function copiedText(...blocks: readonly PMNode[]): string {
  return clipboardText(new Slice(Fragment.from(blocks), 0, 0));
}

/** One paragraph wearing everything the schema keeps to itself */
function loadedParagraph(): PMNode {
  return docxSchema.nodes.paragraph.create(
    {
      srcId: "session:body:3",
      pPr: '<w:pPr><w:pStyle w:val="DocumentHeading1"/></w:pPr>',
      pAttrs: 'w:rsidR="00AB12"',
      format: { align: "center" },
    },
    [
      docxSchema.nodes.commentStart.create({
        id: "0",
        xml: '<w:commentRangeStart w:id="0"/>',
      }),
      docxSchema.text("commented", [
        docxSchema.marks.run.create({
          rPr: "<w:rPr><w:b/></w:rPr>",
          rAttrs: 'w:rsidRPr="00CD34"',
          format: { bold: true },
        }),
      ]),
      docxSchema.nodes.commentEnd.create({
        id: "0",
        xml: '<w:commentRangeEnd w:id="0"/>',
      }),
      docxSchema.nodes.commentReference.create({
        id: "0",
        referenceXml: '<w:commentReference w:id="0"/>',
        author: "Jane Doe",
        authorId: "jane-identity",
        initials: "JD",
        date: "2026-01-01T00:00:00Z",
        paraId: "1A2B3C4D",
        extensionXml: '<w15:commentEx w15:paraId="1A2B3C4D"/>',
        replies: [
          {
            id: "1",
            author: "Bob Roe",
            commentXml: "<w:comment>what the reply says</w:comment>",
          },
        ],
      }),
      docxSchema.nodes.noteReference.create({
        kind: "footnote",
        id: "2",
        label: "1",
        referenceXml: '<w:footnoteReference w:id="2"/>',
      }),
      docxSchema.nodes.rawInline.create({
        xml: '<w:bookmarkStart w:id="7" w:name="secret_bookmark"/>',
        element: "bookmarkStart",
        display: "hidden",
      }),
    ]
  );
}

describe("the outbound clipboard shape", () => {
  it("declares an outbound shape for every node and mark in the schema", () => {
    const nodes = Object.keys(docxSchema.nodes).sort();
    const marks = Object.keys(docxSchema.marks).sort();

    expect(Object.keys(clipboardSpecs.nodes).sort()).toEqual(nodes);
    expect(Object.keys(clipboardSpecs.marks).sort()).toEqual(marks);
    // Every node reaches the serializer through the shape declared for it, so one the schema
    // gained without a shape would be drawn by nothing at all
    expect(() => copiedHtml(loadedParagraph())).not.toThrow();
  });

  it("writes no data- attribute beyond data-style and the channel token, and no raw XML, into the copied HTML", () => {
    const table = docxSchema.nodes.table.create(
      {
        srcId: "session:body:4",
        tblPr: '<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>',
        tblAttrs: 'w:rsidTr="00EF56"',
        gridCols: [1000, 2000],
      },
      docxSchema.nodes.tableRow.create(
        { trPr: '<w:trPr><w:trHeight w:val="400"/></w:trPr>' },
        [
          cell(paragraph(docxSchema.text("A1"))),
          docxSchema.nodes.tableCell.create(
            {
              colspan: 2,
              tcPr: '<w:tcPr><w:shd w:val="clear"/></w:tcPr>',
              sdtPrefix: "<w:sdt><w:sdtPr/>",
            },
            paragraph(docxSchema.text("B1"))
          ),
        ]
      )
    );
    const picture = paragraph(
      docxSchema.nodes.image.create({
        src: TINY_PNG,
        extent: { cx: 952500, cy: 476250 },
        alt: "a seal",
        xml: "<w:drawing>the original drawing</w:drawing>",
      })
    );

    const html = copiedHtml(loadedParagraph(), table, picture);

    expect(html.match(/\sdata-(?!style=)[\w-]+/g)).toBeNull();
    expect(html).not.toContain("<w:");
    expect(html).not.toContain("w15:");
    // The one attribute a copy carries, so another editor can read the style back off it
    expect(html).toContain('data-style="DocumentHeading1"');
    // What the copy is for still travels: the words, the picture, the shape of the table
    expect(html).toContain("commented");
    expect(html).toContain("<td");
    expect(html).toContain('colspan="2"');
    expect(html).toContain(`src="${TINY_PNG}"`);
    expect(html).toContain('alt="a seal"');
  });

  it("leaves comment author, id, text and note bodies out of the copied HTML", () => {
    const html = copiedHtml(loadedParagraph());

    for (const secret of [
      "Jane Doe",
      "jane-identity",
      "JD",
      "2026-01-01T00:00:00Z",
      "1A2B3C4D",
      "what the reply says",
      "secret_bookmark",
      "session:body:3",
    ]) {
      expect(html, secret).not.toContain(secret);
    }
    // The number the reader sees beside the text is not the note, and it does travel
    expect(html).toContain("<sup>1</sup>");
  });

  it("writes tabs, line breaks, page breaks, cell boundaries and note labels into the copied text", () => {
    const lines = paragraph(
      docxSchema.text("before"),
      docxSchema.text("\t", [docxSchema.marks.tab.create()]),
      docxSchema.text("after"),
      docxSchema.nodes.hardBreak.create(),
      docxSchema.text("next line"),
      docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' }),
      docxSchema.text("next page"),
      docxSchema.nodes.noteReference.create({ id: "2", label: "7" })
    );
    const table = docxSchema.nodes.table.create(null, [
      docxSchema.nodes.tableRow.create(null, [
        cell(paragraph(docxSchema.text("one"))),
        cell(paragraph(docxSchema.text("two"))),
      ]),
      docxSchema.nodes.tableRow.create(null, [
        cell(paragraph(docxSchema.text("three"))),
        cell(paragraph(docxSchema.text("four"))),
      ]),
    ]);

    expect(copiedText(lines, table)).toBe(
      "before\tafter\nnext line\fnext page7\none\ttwo\nthree\tfour"
    );
  });

  it("says nothing at all for a comment marker, a bookmark or a note that draws its own mark", () => {
    const marked = paragraph(
      docxSchema.nodes.commentStart.create({ id: "0" }),
      docxSchema.text("body"),
      docxSchema.nodes.commentEnd.create({ id: "0" }),
      docxSchema.nodes.commentReference.create({ id: "0", author: "Jane Doe" }),
      docxSchema.nodes.rawInline.create({
        element: "bookmarkStart",
        display: "hidden",
      }),
      docxSchema.nodes.noteReference.create({
        id: "3",
        label: "2",
        customMarkFollows: true,
      })
    );

    expect(copiedText(marked)).toBe("body");
    expect(copiedHtml(marked)).toBe('<p class="docx-editor-p">body</p>');
  });

  it("carries a kept line and a kept character but not the chip naming a field", () => {
    const kept = paragraph(
      docxSchema.text("first"),
      docxSchema.nodes.rawRunContent.create({
        element: "cr",
        display: "break",
        xml: "<w:cr/>",
      }),
      docxSchema.text("second"),
      docxSchema.nodes.rawRunContent.create({
        element: "noBreakHyphen",
        display: "text",
        text: "‑",
        xml: "<w:noBreakHyphen/>",
      }),
      docxSchema.nodes.rawRunContent.create({
        element: "fldChar",
        display: "chip",
        text: "PAGE",
        xml: '<w:fldChar w:fldCharType="begin"/>',
      })
    );

    expect(copiedText(kept)).toBe("first\nsecond‑");
    const html = copiedHtml(kept);
    expect(html).toContain("<br>");
    expect(html).toContain("<span>‑</span>");
    expect(html).not.toContain("PAGE");
  });

  it("carries a link as an anchor and leaves the relationship it hung off behind", () => {
    const linked = paragraph(
      docxSchema.text("the docs", [
        docxSchema.marks.link.create({
          href: "https://example.com/docs",
          linkPrefix: '<w:hyperlink r:id="RELATIONSHIP">',
          linkKey: 3,
        }),
      ]),
      docxSchema.text(" and ", []),
      docxSchema.text("a script", [
        docxSchema.marks.link.create({ href: "javascript:steal()" }),
      ])
    );

    const html = copiedHtml(linked);

    expect(html).toContain('href="https://example.com/docs"');
    expect(html).not.toContain("RELATIONSHIP");
    // An address no end should act on travels as no address at all
    expect(html).not.toContain("javascript:");
    expect(html).toContain("a script");
  });

  it("names the copy on its first element and nowhere inside it", () => {
    const serializer = clipboardSerializer((copy) => {
      copy.firstElementChild?.setAttribute("data-docx-clip", "the-token");
    });
    const wrap = document.createElement("div");
    wrap.appendChild(
      serializer.serializeFragment(
        Fragment.from([
          paragraph(docxSchema.text("one")),
          paragraph(docxSchema.text("two")),
        ]),
        { document }
      )
    );

    expect(wrap.querySelectorAll("[data-docx-clip]")).toHaveLength(1);
    expect(wrap.firstElementChild?.getAttribute("data-docx-clip")).toBe(
      "the-token"
    );
  });
});
