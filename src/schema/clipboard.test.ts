// @vitest-environment jsdom
import {
  type DOMOutputSpec,
  Fragment,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { DEFAULT_FONT_FALLBACKS } from "../styles/fontStack";
import {
  clipboardSerializer,
  clipboardSpecs,
  clipboardText,
} from "./clipboard";
import { docxSchema } from "./docxSchema";
import { imageNodeSpec, runMarkSpec } from "./rendering";

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

/** The attributes of a drawing that are not this editor talking to itself */
function publicHalf(spec: DOMOutputSpec | null): Record<string, string> {
  const attrs = Array.isArray(spec) ? spec[1] : null;
  if (attrs === null || typeof attrs !== "object") {
    throw new Error("the drawing carries no attributes");
  }
  return Object.fromEntries(
    Object.entries(attrs).flatMap(([name, value]) =>
      name.startsWith("data-") || typeof value !== "string"
        ? []
        : [[name, value]]
    )
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
    // The number the reader sees beside the text is not the note, and it does travel, with what
    // it stands for kept for a reader who is listening to the page rather than seeing it
    expect(html).toContain('<sup aria-label="Footnote 1">1</sup>');
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

  it("carries a kept line and a kept character but not the chip standing for a field character", () => {
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
        text: "\u2011",
        xml: "<w:noBreakHyphen/>",
      }),
      // A field character holds no text of its own; the chip draws the editor's name for it
      docxSchema.nodes.rawRunContent.create({
        element: "fldChar",
        display: "chip",
        xml: '<w:fldChar w:fldCharType="begin"/>',
      })
    );

    expect(copiedText(kept)).toBe("first\nsecond\u2011");
    const html = copiedHtml(kept);
    expect(html).toContain("<br>");
    expect(html).toContain("<span>\u2011</span>");
    expect(html).not.toContain("fldChar");
  });

  /**
   * A chip's text is the text its element held, which is what a reader sees on the page: the
   * result a field last cached, the words a tracked insertion put there.
   */
  it("keeps a cached field result and tracked-change text in the copied HTML and text", () => {
    const kept = paragraph(
      docxSchema.text("page "),
      docxSchema.nodes.rawInline.create({
        element: "fldSimple",
        display: "chip",
        text: "7",
        xml: '<w:fldSimple w:instr=" PAGE ">the instruction</w:fldSimple>',
      }),
      docxSchema.text(" of "),
      docxSchema.nodes.rawInline.create({
        element: "ins",
        display: "chip",
        text: "nine",
        xml: '<w:ins w:author="Jane Doe">the insertion</w:ins>',
      })
    );

    expect(copiedText(kept)).toBe("page 7 of nine");
    const html = copiedHtml(kept);
    expect(html).toContain("<span>7</span>");
    expect(html).toContain("<span>nine</span>");
    // What the fragment is written from stays behind, its author with it
    expect(html).not.toContain("w:instr");
    expect(html).not.toContain("Jane Doe");
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
  it("keeps a cell's own paragraphs and line breaks on one line so a pasted row stays one spreadsheet row", () => {
    const table = docxSchema.nodes.table.create(
      null,
      docxSchema.nodes.tableRow.create(null, [
        cell(
          paragraph(docxSchema.text("top")),
          paragraph(docxSchema.text("bottom"))
        ),
        cell(
          paragraph(
            docxSchema.text("over"),
            docxSchema.nodes.hardBreak.create(),
            docxSchema.text("under")
          )
        ),
        cell(paragraph(docxSchema.text("beside"))),
      ])
    );

    expect(copiedText(table)).toBe("top bottom\tover under\tbeside");
  });

  /**
   * A content control is a wrapper the file keeps, and no application a copy lands in reads one
   * back. The text it held is the part that means anything, and it leaves without the wrapper.
   */
  it("writes the text a content control held without a wrapper standing for the control", () => {
    const controlled = paragraph(
      docxSchema.text("inside the control", [
        docxSchema.marks.sdt.create({
          sdtPrefix: '<w:sdt><w:sdtPr><w:alias w:val="SECRET"/></w:sdtPr>',
          sdtKey: 2,
        }),
      ])
    );

    expect(copiedHtml(controlled)).toBe(
      '<p class="docx-editor-p">inside the control</p>'
    );
    expect(copiedText(controlled)).toBe("inside the control");
  });

  it("opens no blank line where a block the editor only kept stood", () => {
    const kept = docxSchema.nodes.rawBlock.create({
      xml: "<w:tbl>an unsupported table</w:tbl>",
      name: "w:tbl",
      display: "chip",
    });

    expect(
      copiedText(
        paragraph(docxSchema.text("above")),
        kept,
        paragraph(docxSchema.text("below"))
      )
    ).toBe("above\nbelow");
    // On its own it draws nothing but the element the name of the copy rides on
    expect(copiedHtml(kept)).toBe("<span></span>");
  });
  /**
   * The page's drawing and the copy's are written from one pair of functions, so the copy cannot
   * quietly stop saying something the page says, or start saying something private.
   */
  it("copies the public half of the page's own drawing of an image and a run", () => {
    const image = docxSchema.nodes.image.create({
      src: TINY_PNG,
      extent: { cx: 952500, cy: 476250 },
      alt: "a seal",
      xml: "<w:drawing>the original drawing</w:drawing>",
    });
    const run = docxSchema.marks.run.create({
      rPr: "<w:rPr><w:b/></w:rPr>",
      format: { bold: true, lang: "ko-KR" },
    });
    const drawImage = clipboardSpecs.nodes.image.toClipboardDOM;
    const drawRun = clipboardSpecs.marks.run.toClipboardDOM;
    if (drawImage === null || drawRun === null) {
      throw new Error("an image and a run both leave the editor");
    }

    expect(publicHalf(drawImage(image))).toEqual(
      publicHalf(imageNodeSpec(image.attrs))
    );
    expect(publicHalf(drawRun(run))).toEqual(
      publicHalf(runMarkSpec(run.attrs, DEFAULT_FONT_FALLBACKS))
    );
    // And the page keeps saying the things a copy does not
    expect(imageNodeSpec(image.attrs)).toContainEqual(
      expect.objectContaining({
        "data-xml": "<w:drawing>the original drawing</w:drawing>",
      })
    );
  });
});
