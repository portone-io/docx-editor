// @vitest-environment jsdom
import {
  DOMSerializer,
  Fragment,
  DOMParser as PMDOMParser,
  type Node as PMNode,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  toCellFormat,
  toParagraphFormat,
  toRunFormat,
  toTableFormat,
} from "../model/format";
import { editorClassNames } from "../styles/classNames";
import { docxSchema } from "./index";

const serializer = DOMSerializer.fromSchema(docxSchema);
const parser = PMDOMParser.fromSchema(docxSchema);

function render(...blocks: PMNode[]): HTMLElement {
  const host = document.createElement("div");
  host.appendChild(serializer.serializeFragment(Fragment.fromArray(blocks)));
  return host;
}

function paragraph(attrs: Record<string, unknown>, inline: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(attrs, inline);
}

function run(text: string, attrs: Record<string, unknown>): PMNode {
  return docxSchema.text(text, [docxSchema.marks.run.create(attrs)]);
}

function cell(attrs: Record<string, unknown>, text: string): PMNode {
  return docxSchema.nodes.tableCell.create(attrs, [
    paragraph({}, text ? [docxSchema.text(text)] : []),
  ]);
}

function tableRow(...cells: PMNode[]): PMNode {
  return docxSchema.nodes.tableRow.create(null, cells);
}

function table(attrs: Record<string, unknown>, rows: PMNode[]): PMNode {
  return docxSchema.nodes.table.create(attrs, rows);
}

/** A 1x1 png, the shortest src an image node can carry */
const PNG_SRC =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAf" +
  "FcSJAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=";

describe("toDOM", () => {
  it("renders a bold run as font-weight", () => {
    const html = render(
      paragraph({}, [run("bold", { format: { bold: true } })])
    ).innerHTML;
    expect(html).toContain("font-weight: bold");
  });

  it("renders the font size and the highlight", () => {
    const html = render(
      paragraph({}, [
        run("emphasis", { format: { fontSizePt: 11, highlight: "yellow" } }),
      ])
    ).innerHTML;
    expect(html).toContain("font-size: 11pt");
    expect(html).toContain("background-color: rgb(255, 255, 0)");
  });

  it("renders the language of a run as the lang of its span", () => {
    const host = render(
      paragraph({}, [run("clause", { format: { lang: "ja-JP" } })])
    );
    expect(host.querySelector("span")?.getAttribute("lang")).toBe("ja-JP");
  });

  it("draws no lang at all for a run that records no language", () => {
    const host = render(paragraph({}, [run("clause", { format: {} })]));
    expect(host.querySelector("span")?.hasAttribute("lang")).toBe(false);
    // A value that is not a language tag never reaches the attribute
    const tampered = render(
      paragraph({}, [run("clause", { format: { lang: 'ja" onload="x' } })])
    );
    expect(tampered.querySelector("span")?.hasAttribute("lang")).toBe(false);
  });

  it("renders paragraph alignment and spacing as styles", () => {
    const html = render(
      paragraph({ format: { align: "center", spaceAfterPt: 10 } }, [])
    ).innerHTML;
    expect(html).toContain("text-align: center");
    expect(html).toContain("margin-bottom: 10pt");
  });

  /**
   * Text typed in the editor carries no run of its own, so the character formatting of the style
   * the paragraph wears can only reach it by inheritance
   */
  it("renders the character formatting the paragraph's style lays down", () => {
    const html = render(
      paragraph(
        {
          styleRun: {
            italic: true,
            color: "#2e74b5",
            fontSizePt: 16,
            background: "#ffff00",
          },
        },
        [docxSchema.text("draft")]
      )
    ).innerHTML;
    expect(html).toContain("font-style: italic");
    expect(html).toContain("color: rgb(46, 116, 181)");
    expect(html).toContain("font-size: 16pt");
    // A fill paints a box instead of carrying down, so the paragraph does not take that one on
    expect(html).not.toContain("background-color");
  });

  it("renders a table it could not model as a placeholder box naming it", () => {
    const host = render(
      docxSchema.nodes.rawBlock.create({
        srcId: "opened:body:1",
        name: "w:tbl",
      })
    );
    expect(host.textContent).toContain("Table (unsupported layout");
  });

  it("renders a block of another kind as a placeholder that does not claim to be a table", () => {
    const host = render(
      docxSchema.nodes.rawBlock.create({
        srcId: "opened:body:1",
        name: "w:sectPr",
      })
    );
    expect(host.textContent).toContain("Unsupported content");
  });

  it("renders a body-level bookmark as a hidden marker without placeholder text", () => {
    const host = render(
      docxSchema.nodes.rawBlock.create({
        srcId: "opened:body:1",
        name: "w:bookmarkStart",
        display: "hidden",
        guarded: true,
      })
    );
    const marker = host.querySelector(".docx-editor-raw-block");

    expect(marker?.getAttribute("data-display")).toBe("hidden");
    expect(host.textContent).toBe("");
  });

  it("renders a table's grid widths as a colgroup", () => {
    const host = render(
      table({ gridCols: [1500, 3000], tblW: { type: "dxa", twips: 4500 } }, [
        tableRow(cell({}, "a"), cell({}, "b")),
      ])
    );
    const cols = Array.from(host.querySelectorAll("col"));
    expect(cols.map((col) => col.getAttribute("style"))).toEqual([
      "width: 100px;",
      "width: 200px;",
    ]);
    expect(host.querySelector("table")?.getAttribute("style")).toContain(
      "width: 300px"
    );
  });

  it("renders a percentage width table as a percentage", () => {
    // w:tblW w:type="pct" w:w="5000" means 100%. Misreading this value folds the table up into a strip
    const host = render(
      table({ gridCols: [100, 100], tblW: { type: "pct", fiftieths: 5000 } }, [
        tableRow(cell({}, "a"), cell({}, "b")),
      ])
    );
    expect(host.querySelector("table")?.getAttribute("style")).toContain(
      "width: 100%"
    );
  });

  it("renders merged cells as colspan and rowspan", () => {
    const host = render(
      table({ gridCols: [1000, 1000, 1000] }, [
        tableRow(
          cell({ rowspan: 2 }, "PartyB"),
          cell({ colspan: 2 }, "Company")
        ),
        tableRow(cell({}, "a"), cell({}, "b")),
      ])
    );
    const cells = Array.from(host.querySelectorAll("td"));
    expect(
      cells.map((td) => [
        td.getAttribute("colspan"),
        td.getAttribute("rowspan"),
      ])
    ).toEqual([
      [null, "2"],
      ["2", null],
      [null, null],
      [null, null],
    ]);
  });

  it("renders cell formatting as styles", () => {
    const host = render(
      table({}, [
        tableRow(
          cell(
            {
              format: {
                borderTop: "0.5pt solid #000000",
                background: "#bfbfbf",
                verticalAlign: "center",
                paddingLeftPt: 5,
              },
            },
            "a"
          )
        ),
      ])
    );
    const style = host.querySelector("td")?.getAttribute("style") ?? "";
    expect(style).toContain("border-top: 0.5pt solid rgb(0, 0, 0)");
    expect(style).toContain("background-color: rgb(191, 191, 191)");
    expect(style).toContain("vertical-align: middle");
    expect(style).toContain("padding-left: 5pt");
  });

  it("renders row height as a style", () => {
    const host = render(
      table({}, [
        docxSchema.nodes.tableRow.create(
          { format: { height: { rule: "atLeast", pt: 27 } } },
          [cell({}, "a")]
        ),
      ])
    );
    expect(host.querySelector("tr")?.getAttribute("style")).toContain(
      "height: 27pt"
    );
  });

  it("a block inside a cell that could not be modeled stays as a placeholder", () => {
    const host = render(
      table({}, [
        tableRow(
          docxSchema.nodes.tableCell.create({}, [
            docxSchema.nodes.rawBlock.create({
              xml: "<w:tbl/>",
              name: "w:tbl",
            }),
          ])
        ),
      ])
    );
    expect(host.innerHTML).toContain("docx-editor-raw-block");
    expect(host.textContent).toContain("original is preserved");
  });

  it("a bookmark stays as a spot with no width", () => {
    const host = render(
      paragraph({}, [
        docxSchema.nodes.rawInline.create({
          xml: '<w:bookmarkStart w:id="0" w:name="here"/>',
        }),
      ])
    );
    expect(host.innerHTML).toContain("docx-editor-raw-inline");
    expect(host.textContent).toBe("");
  });

  it("renders an image at the size the drawing recorded", () => {
    const host = render(
      paragraph({}, [
        docxSchema.nodes.image.create({
          src: PNG_SRC,
          extent: { cx: 1905000, cy: 952500 },
          alt: "a seal",
          xml: "<w:drawing/>",
        }),
      ])
    );
    const img = host.querySelector("img");
    expect(img?.className).toBe("docx-editor-img");
    expect(img?.getAttribute("src")).toBe(PNG_SRC);
    expect(img?.getAttribute("alt")).toBe("a seal");
    // 9525 units to the pixel
    expect(img?.getAttribute("width")).toBe("200");
    expect(img?.getAttribute("height")).toBe("100");
  });

  it("only a page breaking line break comes out carrying a marker", () => {
    const host = render(
      paragraph({}, [
        docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' }),
        docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="textWrapping"' }),
        docxSchema.nodes.hardBreak.create({ brAttrs: null }),
      ])
    );
    const breaks = Array.from(host.querySelectorAll("br"));
    expect(breaks.map((br) => br.getAttribute("data-break"))).toEqual([
      "page",
      null,
      null,
    ]);
  });

  it("marks a paragraph that starts a new page from here", () => {
    const host = render(
      paragraph({ format: { pageBreakBefore: true } }, []),
      paragraph({ format: { align: "left" } }, [])
    );
    expect(host.querySelectorAll("p[data-page-break]").length).toBe(1);
  });
});

describe("parseDOM", () => {
  const original = docxSchema.nodes.doc.create(null, [
    paragraph(
      {
        srcId: "opened:body:0",
        pAttrs: 'w:rsidR="00A"',
        pPr: '<w:pPr><w:keepNext/><w:pageBreakBefore/><w:jc w:val="center"/></w:pPr>',
        format: { align: "center", keepNext: true, pageBreakBefore: true },
        styleRun: { italic: true, fontSizePt: 16 },
      },
      [
        run("  padded text  ", {
          rPr: '<w:rPr><w:b/><w:rtl w:val="0"/></w:rPr>',
          rAttrs: 'w:rsidRPr="00B"',
          format: { bold: true },
        }),
        docxSchema.text("\t", [docxSchema.marks.tab.create()]),
        docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' }),
        docxSchema.nodes.rawInline.create({ xml: '<w:bookmarkEnd w:id="0"/>' }),
        docxSchema.nodes.image.create({
          src: PNG_SRC,
          extent: { cx: 1905000, cy: 952500 },
          alt: "a seal",
          xml: '<w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/></wp:inline></w:drawing>',
        }),
      ]
    ),
    table(
      {
        srcId: "opened:body:1",
        tblAttrs: 'w:rsidR="00C"',
        tblPr: '<w:tblPr><w:tblW w:type="dxa" w:w="4500"/></w:tblPr>',
        tblW: { type: "dxa", twips: 4500 },
        gridCols: [1500, 1500, 1500],
        format: { borderTop: "0.5pt solid #000000", align: "center" },
        styleInside: { horizontal: "0.5pt solid #999999", vertical: null },
        styleCellMargins: {
          topPt: 0,
          rightPt: 5.4,
          bottomPt: 0,
          leftPt: 5.4,
        },
        styleConditions: {
          firstRow: {
            background: "#D9E2F3",
            borders: {
              top: null,
              bottom: "1pt solid #000000",
              left: null,
              right: null,
            },
            inside: { horizontal: null, vertical: "none" },
            margins: {
              topPt: null,
              rightPt: null,
              bottomPt: null,
              leftPt: null,
            },
            verticalAlign: "center",
          },
        },
        styleBands: { row: 2, col: null },
      },
      [
        docxSchema.nodes.tableRow.create(
          {
            trAttrs: 'w:rsidR="00D"',
            trPr: '<w:trPr><w:trHeight w:val="540" w:hRule="atLeast"/></w:trPr>',
            format: { height: { rule: "atLeast", pt: 27 } },
          },
          [
            cell(
              {
                rowspan: 2,
                tcAttrs: 'w:rsidR="00E"',
                tcPr:
                  '<w:tcPr><w:tcW w:type="dxa" w:w="1500"/>' +
                  '<w:vMerge w:val="restart"/></w:tcPr>',
                tcW: { type: "dxa", twips: 1500 },
                format: { verticalAlign: "center", paddingLeftPt: 5 },
              },
              "PartyB"
            ),
            cell({ colspan: 2, colwidth: [100, 200] }, "Company"),
          ]
        ),
        tableRow(cell({}, "a"), cell({}, "b")),
      ]
    ),
    docxSchema.nodes.rawBlock.create({ srcId: "opened:body:2", name: "w:tbl" }),
    docxSchema.nodes.rawBlock.create({
      srcId: "opened:body:3",
      name: "w:sectPr",
    }),
  ]);

  it("reading the rendered DOM back loses neither the original XML nor the display formatting", () => {
    const host = render(...original.children);
    const reparsed = parser.parse(host, { preserveWhitespace: true });
    expect(reparsed.eq(original)).toBe(true);
  });

  /**
   * If a `div[data-src]` that came in from outside became a preserved block, a node
   * pointing at an original fragment that does not exist would settle into the
   * document.
   */
  it("a div[data-src] we did not render does not become a preserved block", () => {
    const host = document.createElement("div");
    host.innerHTML = '<div data-src="0" data-name="w:tbl">outer text</div>';
    const parsed = parser.parse(host);

    expect(parsed.firstChild?.type.name).not.toBe("rawBlock");
    expect(parsed.textContent).toBe("outer text");
  });

  /**
   * A block with no original and no name behind it still needs both attributes, because what
   * recognizes it is the class and `data-name` together. Reading either empty string back as a
   * value would point the block at a fragment or an element that is not its own.
   */
  it("a data-src and a data-name that are empty read back as neither", () => {
    const host = render(docxSchema.nodes.rawBlock.create({ xml: "<w:tbl/>" }));
    const parsed = parser.parse(host);

    expect(parsed.firstChild?.type.name).toBe("rawBlock");
    expect(parsed.firstChild?.attrs).toMatchObject({
      srcId: null,
      name: null,
    });
  });

  it("a stretch of text inside a content control round trips too", () => {
    const control = docxSchema.marks.sdt.create({
      sdtPrefix: '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr>',
      contentsLocked: true,
      deletionLocked: true,
    });
    const withControl = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        run("signed on ", {}),
        docxSchema.text("2026-08-04", [
          control,
          docxSchema.marks.run.create({}),
        ]),
        docxSchema.nodes.rawInline.create(
          { xml: "<w:r><w:rPr/></w:r>" },
          null,
          [control]
        ),
      ]),
    ]);
    const host = render(...withControl.children);
    // Both nodes wear the same mark, so one single span is drawn around the two of them
    expect(host.querySelectorAll(".docx-editor-sdt-locked").length).toBe(1);
    expect(
      parser.parse(host, { preserveWhitespace: true }).eq(withControl)
    ).toBe(true);
  });

  /**
   * Two controls whose opening XML is identical are told apart by their number alone. Were that
   * number missing from the DOM, the text of one would run into the next on the way back in
   * (a paste, or a redraw of the whole document).
   */
  it("two neighbouring content controls stay apart through the DOM", () => {
    const controlAt = (key: number) =>
      docxSchema.marks.sdt.create({
        sdtPrefix: '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>',
        key,
      });
    const neighbours = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        docxSchema.text("PartyA", [controlAt(0)]),
        docxSchema.text("PartyB", [controlAt(1)]),
      ]),
    ]);
    const host = render(...neighbours.children);

    expect(
      Array.from(host.querySelectorAll(".docx-editor-sdt")).map((span) =>
        span.getAttribute("data-key")
      )
    ).toEqual(["0", "1"]);
    expect(
      parser.parse(host, { preserveWhitespace: true }).eq(neighbours)
    ).toBe(true);
  });

  it("a stretch of text inside a hyperlink round trips too", () => {
    const link = docxSchema.marks.link.create({
      linkPrefix: '<w:hyperlink r:id="rId9" w:tooltip="Our terms">',
      href: "https://example.com/terms",
    });
    const withLink = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        run("see ", {}),
        docxSchema.text("our terms", [link, docxSchema.marks.run.create({})]),
      ]),
    ]);
    const host = render(...withLink.children);

    // The address is drawn as data, not as an href: nothing here is a link the browser follows
    const drawn = host.querySelector(".docx-editor-link");
    expect(drawn?.getAttribute("data-href")).toBe("https://example.com/terms");
    expect(host.querySelector("a")).toBeNull();
    expect(parser.parse(host, { preserveWhitespace: true }).eq(withLink)).toBe(
      true
    );
  });

  it("marks an underlined run so a link over it draws no second line", () => {
    const link = docxSchema.marks.link.create({
      linkPrefix: '<w:hyperlink r:id="rId9">',
      href: "https://example.com/terms",
    });
    const styled = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        docxSchema.text("our terms", [
          link,
          docxSchema.marks.run.create({ format: { underline: "single" } }),
        ]),
        run("plainly", { format: { underline: "none" } }),
      ]),
    ]);
    const host = render(...styled.children);

    expect(
      Array.from(host.querySelectorAll(`.${editorClassNames.run}`)).map(
        (span) => span.hasAttribute("data-underline")
      )
    ).toEqual([true, false]);
    expect(parser.parse(host, { preserveWhitespace: true }).eq(styled)).toBe(
      true
    );
  });

  /** Two links written exactly alike are told apart by their number alone, as two controls are */
  it("two neighbouring hyperlinks stay apart through the DOM", () => {
    const linkAt = (key: number) =>
      docxSchema.marks.link.create({
        linkPrefix: '<w:hyperlink r:id="rId9">',
        href: "https://example.com",
        key,
      });
    const neighbours = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        docxSchema.text("terms", [linkAt(0)]),
        docxSchema.text("prices", [linkAt(1)]),
      ]),
    ]);
    const host = render(...neighbours.children);

    expect(
      Array.from(host.querySelectorAll(".docx-editor-link")).map((span) =>
        span.getAttribute("data-key")
      )
    ).toEqual(["0", "1"]);
    expect(
      parser.parse(host, { preserveWhitespace: true }).eq(neighbours)
    ).toBe(true);
  });

  it("a link the editor made, which has no opening tag of its own, round trips too", () => {
    const made = docxSchema.nodes.doc.create(null, [
      paragraph({}, [
        docxSchema.text("terms", [
          docxSchema.marks.link.create({ href: "https://example.com" }),
        ]),
      ]),
    ]);
    const host = render(...made.children);
    expect(parser.parse(host, { preserveWhitespace: true }).eq(made)).toBe(
      true
    );
  });

  it("a cell wrapped in a locked control round trips too", () => {
    const withLockedCell = docxSchema.nodes.doc.create(null, [
      table({ gridCols: [1000] }, [
        tableRow(
          cell(
            {
              sdtPrefix:
                '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr>',
              sdtContentsLocked: true,
              sdtDeletionLocked: true,
            },
            "final"
          )
        ),
      ]),
    ]);
    const host = render(...withLockedCell.children);

    expect(host.querySelectorAll("td.docx-editor-tc-locked").length).toBe(1);
    expect(
      parser.parse(host, { preserveWhitespace: true }).eq(withLockedCell)
    ).toBe(true);
  });

  it("a preserved block inside a cell round trips too", () => {
    const withRaw = docxSchema.nodes.doc.create(null, [
      table({ gridCols: [1000] }, [
        tableRow(
          docxSchema.nodes.tableCell.create({}, [
            paragraph({}, [docxSchema.text("outer")]),
            docxSchema.nodes.rawBlock.create({
              xml: "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
              name: "w:tbl",
            }),
          ])
        ),
      ]),
    ]);
    const host = render(...withRaw.children);
    expect(parser.parse(host, { preserveWhitespace: true }).eq(withRaw)).toBe(
      true
    );
  });
});

/**
 * A preserved fragment goes back out spliced into a slot the writer opens and closes around it.
 * One that closes that slot itself, or opens a sibling beside it, would write content into the
 * exported file that never stood in the document, so a rule reading one turns it down instead.
 */
describe("raw XML coming in through the DOM", () => {
  function parseHtml(html: string): PMNode {
    const host = document.createElement("div");
    host.innerHTML = html;
    return parser.parse(host);
  }

  it("a paragraph whose data-ppr closes the paragraph is read as an unstyled paragraph", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}" data-ppr="&lt;/w:p&gt;&lt;w:p&gt;&lt;w:pPr&gt;&lt;w:jc w:val=&quot;center&quot;/&gt;&lt;/w:pPr&gt;">x</p>`
    );

    expect(parsed.firstChild?.attrs.pPr).toBeNull();
    expect(parsed.textContent).toBe("x");
  });

  it("a paragraph whose data-pattrs closes the opening tag is read as an unstyled paragraph", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}" data-pattrs="w:rsidR=&quot;00A&quot;&gt;&lt;w:r&gt;&lt;w:t&gt;smuggled&lt;/w:t&gt;&lt;/w:r">x</p>`
    );

    expect(parsed.firstChild?.attrs.pAttrs).toBeNull();
    expect(parsed.textContent).toBe("x");
  });

  it("keeps a paragraph whose only attribute is a w14 paraId", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}" data-pattrs="w14:paraId=&quot;1A2B3C4D&quot;">x</p>`
    );

    expect(parsed.firstChild?.attrs.pAttrs).toBe('w14:paraId="1A2B3C4D"');
  });

  /**
   * A producer is free to declare a namespace on the paragraph that uses it rather than on the
   * root of the part, and .NET's `XmlWriter` does, so turning one down would drop the properties
   * of every paragraph in such a file on the first re-read of the live DOM.
   */
  it("keeps a paragraph that declares the namespace its own attribute uses", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}" data-pattrs="xmlns:w14=&quot;u&quot; w14:paraId=&quot;1&quot;">x</p>`
    );

    expect(parsed.firstChild?.attrs.pAttrs).toBe(
      'xmlns:w14="u" w14:paraId="1"'
    );
  });

  it("a paragraph whose data-pattrs rebinds the relationship namespace is read as an unstyled paragraph", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}" data-pattrs="xmlns:r=&quot;urn:evil&quot;">x</p>`
    );

    expect(parsed.firstChild?.attrs.pAttrs).toBeNull();
    expect(parsed.textContent).toBe("x");
  });

  it("a run whose data-rpr smuggles a w:t loses the mark and keeps the text", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.run}" data-rpr="&lt;w:rPr&gt;&lt;w:b/&gt;&lt;/w:rPr&gt;&lt;w:t&gt;smuggled&lt;/w:t&gt;">x</span></p>`
    );

    expect(parsed.textContent).toBe("x");
    expect(parsed.firstChild?.firstChild?.marks).toEqual([]);
  });

  it("a break whose data-battrs opens a child of its own is not read as a break", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}">a<br data-battrs="w:type=&quot;page&quot;&gt;&lt;w:t&gt;smuggled&lt;/w:t&gt;">b</p>`
    );

    expect(parsed.firstChild?.childCount).toBe(1);
    expect(parsed.firstChild?.child(0).type.name).toBe("text");
  });

  it("a rawInline carrying a comment range marker beside its element is refused", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.rawInline}" data-xml="&lt;w:bookmarkEnd w:id=&quot;1&quot;/&gt;&lt;w:commentRangeStart w:id=&quot;7&quot;/&gt;"></span>x</p>`
    );

    expect(parsed.firstChild?.childCount).toBe(1);
    expect(parsed.textContent).toBe("x");
  });

  it("a rawInline keeps an element of a namespace the editor does not model", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.rawInline}" data-xml="&lt;m:oMathPara/&gt;"></span></p>`
    );

    expect(parsed.firstChild?.firstChild?.attrs.xml).toBe("<m:oMathPara/>");
  });

  it("a comment marker whose data-xml is not a range marker is refused", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.commentMarker}" data-comment-marker="start" data-comment-id="7" data-xml="&lt;w:p&gt;&lt;w:r&gt;&lt;w:t&gt;smuggled&lt;/w:t&gt;&lt;/w:r&gt;&lt;/w:p&gt;"></span>x</p>`
    );

    expect(parsed.firstChild?.childCount).toBe(1);
    expect(parsed.textContent).toBe("x");
  });

  it("a comment reference whose reply carries smuggled thread state is refused", () => {
    const replies = JSON.stringify([
      {
        id: "2",
        paraId: "0A0A0A0A",
        parentParaId: "0B0B0B0B",
        extensionXml:
          '<w15:commentEx w15:paraId="0A0A0A0A"/><w15:commentEx w15:paraId="0C0C0C0C"/>',
      },
    ]);
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.commentMarker}" data-comment-marker="reference" data-comment-id="1" data-comment-replies="${replies.replaceAll('"', "&quot;")}"></span>x</p>`
    );

    expect(parsed.firstChild?.childCount).toBe(1);
    expect(parsed.textContent).toBe("x");
  });

  it("a note reference whose data-reference-xml names another element is refused", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><sup class="${editorClassNames.noteReference}" data-note-id="1" data-reference-xml="&lt;w:commentReference w:id=&quot;1&quot;/&gt;">1</sup></p>`
    );

    expect(parsed.firstChild?.child(0).type.name).not.toBe("noteReference");
  });

  it("an image whose data-xml is not a drawing is refused", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><img class="${editorClassNames.image}" src="data:image/png;base64,iVBORw0KGgo=" data-xml="&lt;w:drawing/&gt;&lt;w:t&gt;smuggled&lt;/w:t&gt;"></p>`
    );

    expect(parsed.firstChild?.childCount).toBe(0);
  });

  it("a preserved block whose data-xml opens a sibling is refused", () => {
    const parsed = parseHtml(
      `<div class="${editorClassNames.rawBlock} ${editorClassNames.rawBlock}" data-xml="&lt;w:tbl/&gt;&lt;w:tbl/&gt;" data-name="w:tbl">placeholder</div>`
    );

    expect(parsed.firstChild?.type.name).not.toBe("rawBlock");
  });

  it("a cell whose data-tcpr closes the cell is read as a cell without properties", () => {
    const parsed = parseHtml(
      `<table class="${editorClassNames.table}"><tbody><tr><td class="${editorClassNames.tableCell}" data-tcpr="&lt;/w:tc&gt;&lt;w:tc&gt;&lt;w:tcPr/&gt;"><p class="${editorClassNames.paragraph}">x</p></td></tr></tbody></table>`
    );

    const cell = parsed.firstChild?.firstChild?.firstChild;
    expect(cell?.type.name).toBe("tableCell");
    expect(cell?.attrs.tcPr).toBeNull();
    expect(parsed.textContent).toBe("x");
  });

  it("a control whose opening tag names another element loses the mark", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.sdt}" data-sdt-prefix="&lt;w:tbl&gt;">x</span></p>`
    );

    expect(parsed.textContent).toBe("x");
    expect(parsed.firstChild?.firstChild?.marks).toEqual([]);
  });

  it("a link whose opening tag does not close keeps neither the tag nor the address", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.link}" data-link-prefix="&lt;w:hyperlink&gt;&lt;w:r&gt;" data-href="https://example.com">x</span></p>`
    );

    expect(parsed.textContent).toBe("x");
    expect(parsed.firstChild?.firstChild?.marks).toEqual([]);
  });

  it("a tab whose data-tattrs is not an attribute list loses the mark", () => {
    const parsed = parseHtml(
      `<p class="${editorClassNames.paragraph}"><span class="${editorClassNames.tab}" data-tattrs="&lt;w:tab/&gt;">\t</span></p>`
    );

    expect(parsed.firstChild?.firstChild?.marks).toEqual([]);
  });
});

/**
 * Display values go out to the screen as `data-fmt` and come back in again.
 * A value tampered with in between must never flow into a CSS declaration.
 */
describe("display values coming in through data-fmt", () => {
  it("accepts only six hex digits for a color", () => {
    expect(toRunFormat({ color: "#1155cc" })?.color).toBe("#1155cc");
    expect(toRunFormat({ color: "red;position:fixed" })?.color).toBeUndefined();
    expect(
      toRunFormat({ background: "url(evil)" })?.background
    ).toBeUndefined();
    expect(toCellFormat({ background: "#bfbfbf" })?.background).toBe("#bfbfbf");
  });

  it("accepts only a list of quoted names for a font", () => {
    const names = '"Malgun Gothic","Arial"';
    expect(toRunFormat({ fontFamily: names })?.fontFamily).toBe(names);
    expect(toRunFormat({ fontFamily: "Arial" })?.fontFamily).toBeUndefined();
    expect(
      toRunFormat({ fontFamily: '"a";position:fixed;x:"b"' })?.fontFamily
    ).toBeUndefined();
  });

  it("accepts only known names for underline and highlight", () => {
    expect(toRunFormat({ underline: "single" })?.underline).toBe("single");
    expect(toRunFormat({ underline: "solid" })?.underline).toBeUndefined();
    expect(toRunFormat({ highlight: "yellow" })?.highlight).toBe("yellow");
    expect(toRunFormat({ highlight: "#ffff00" })?.highlight).toBeUndefined();
  });

  it("accepts only the shape the reader produces for a border", () => {
    const drawn = "0.5pt solid #000000";
    expect(toCellFormat({ borderTop: drawn })?.borderTop).toBe(drawn);
    // A cell also carries "do not draw this edge" as a display value
    expect(toCellFormat({ borderBottom: "none" })?.borderBottom).toBe("none");
    expect(toCellFormat({ borderLeft: "1pt solid red" })?.borderLeft).toBe(
      undefined
    );
    expect(
      toParagraphFormat({ borderRight: `${drawn};position:fixed` })?.borderRight
    ).toBeUndefined();
    expect(
      toTableFormat({ borderTop: "1pt solid #000000 inset" })?.borderTop
    ).toBeUndefined();
  });

  it("accepts only a data URL of a drawable kind for an image", () => {
    const host = document.createElement("div");
    host.innerHTML =
      '<p class="docx-editor-p"><img class="docx-editor-img" src="javascript:alert(1)"' +
      ' data-extent=\'{"cx":100,"cy":100}\'></p>';
    const parsed = parser.parse(host, { preserveWhitespace: true });
    const image = parsed.child(0).child(0);

    expect(image.type.name).toBe("image");
    expect(image.attrs.src).toBe(null);
    expect(render(parsed.child(0)).innerHTML).not.toContain("javascript:");
  });

  it("a tampered data-fmt never reaches the on screen style", () => {
    const host = document.createElement("div");
    host.innerHTML =
      '<p class="docx-editor-p" data-fmt=\'{"background":"red;position:fixed"}\'>text</p>';
    const parsed = parser.parse(host, { preserveWhitespace: true });

    expect(toParagraphFormat(parsed.child(0).attrs.format)).toEqual({});
    expect(render(parsed.child(0)).innerHTML).not.toContain("fixed");
  });
});

describe("the two kinds of block", () => {
  it("every block names exactly one of modelled and preserved", () => {
    const blocks = Object.values(docxSchema.nodes).filter((type) =>
      type.isInGroup("block")
    );
    expect(blocks.length).toBeGreaterThan(0);
    for (const type of blocks) {
      const kinds = ["modelled", "preserved"].filter((group) =>
        type.isInGroup(group)
      );
      expect([type.name, kinds]).toEqual([type.name, [expect.any(String)]]);
    }
  });
});
