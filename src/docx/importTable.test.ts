// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { fixTables } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { fixtureNames, readFixture } from "../__testing__/docx";
import { parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { readStyles } from "./formatting";
import { importDocx } from "./importDocx";
import { NO_IMPORT_SOURCES } from "./importParagraph";
import { buildTable } from "./importTable";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** The single element a fragment holds */
function element(xml: string): Element {
  const wrapped = parseXml(`<w:wrap ${W_NS}>${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return el;
}

/** Moves a single XML fragment into a table node */
function table(xml: string): PMNode | null {
  return buildTable(element(xml), 0);
}

function requireTable(xml: string): PMNode {
  const node = table(xml);
  if (!node) throw new Error("the table could not be modelled");
  return node;
}

/** Moves a fragment into a table node with the styles of a styles.xml underneath it */
function styledTable(
  xml: string,
  styles: string,
  defaultTableStyleId: string | null = null
): PMNode {
  const node = buildTable(
    element(xml),
    0,
    NO_IMPORT_SOURCES,
    readStyles(parseXml(`<w:styles ${W_NS}>${styles}</w:styles>`)),
    defaultTableStyleId
  );
  if (!node) throw new Error("the table could not be modelled");
  return node;
}

/** A picture that records the `horizontal x vertical` merge counts for each row */
function spans(node: PMNode): string[] {
  const rows: string[] = [];
  node.forEach((row) => {
    const cells: string[] = [];
    row.forEach((cell) =>
      cells.push(`${cell.attrs.colspan}x${cell.attrs.rowspan}`)
    );
    rows.push(cells.join(" "));
  });
  return rows;
}

function texts(node: PMNode): string[] {
  const rows: string[] = [];
  node.forEach((row) => {
    const cells: string[] = [];
    row.forEach((cell) => cells.push(cell.textContent));
    rows.push(cells.join("|"));
  });
  return rows;
}

const grid = (...cols: number[]) =>
  `<w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;

/** A single `<w:tc>`. Takes a tcPr fragment and some text */
const cell = (tcPr: string, text = "") =>
  `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ""}` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const row = (...cells: string[]) => `<w:tr>${cells.join("")}</w:tr>`;

describe("gridSpan", () => {
  it("moves the horizontal merge count over to colspan", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000, 1000) +
        row(cell('<w:gridSpan w:val="2"/>', "wide cell"), cell("", "single")) +
        row(cell("", "a"), cell("", "b"), cell("", "c")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["2x1 1x1", "1x1 1x1 1x1"]);
    expect(texts(node)).toEqual(["wide cell|single", "a|b|c"]);
  });
});

describe("vMerge", () => {
  it("counts it as the rowspan of the restart cell and builds no continuation cells", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(
          cell('<w:vMerge w:val="restart"/>', "PartyB"),
          cell("", "Company")
        ) +
        row(cell("<w:vMerge/>"), cell("", "Address")) +
        row(cell("<w:vMerge/>"), cell("", "Contact")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["1x3 1x1", "1x1", "1x1"]);
    expect(texts(node)).toEqual(["PartyB|Company", "Address", "Contact"]);
  });

  it("a val of continue is a continuation cell too", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(cell('<w:vMerge w:val="restart"/>', "top"), cell("", "a")) +
        row(cell('<w:vMerge w:val="continue"/>'), cell("", "b")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["1x2 1x1", "1x1"]);
  });

  it("a merge can break and start again in the same column", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(cell('<w:vMerge w:val="restart"/>', "first"), cell("", "a")) +
        row(cell("<w:vMerge/>"), cell("", "b")) +
        row(cell('<w:vMerge w:val="restart"/>', "second"), cell("", "c")) +
        row(cell("<w:vMerge/>"), cell("", "d")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["1x2 1x1", "1x1", "1x2 1x1", "1x1"]);
    expect(texts(node)).toEqual(["first|a", "b", "second|c", "d"]);
  });

  it("a horizontal merge and a vertical merge can overlap on one cell", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000, 1000) +
        row(
          cell(
            '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>',
            "wide and tall"
          ),
          cell("", "a")
        ) +
        row(cell('<w:gridSpan w:val="2"/><w:vMerge/>'), cell("", "b")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["2x2 1x1", "1x1"]);
  });
});

describe("a row that departs from the table's own properties", () => {
  const TBL_PR_EX =
    '<w:tblPrEx><w:tblBorders><w:top w:val="none"/></w:tblBorders></w:tblPrEx>';

  it("stays editable and holds on to the tblPrEx as it is", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000) +
        `<w:tr>${TBL_PR_EX}<w:trPr><w:cantSplit w:val="0"/></w:trPr>` +
        cell("", "a") +
        "</w:tr></w:tbl>"
    );
    expect(texts(node)).toEqual(["a"]);
    expect(node.child(0).attrs.tblPrEx).toBe(TBL_PR_EX);
    expect(node.child(0).attrs.trPr).toBe(
      '<w:trPr><w:cantSplit w:val="0"/></w:trPr>'
    );
  });

  it("a row without one carries no fragment", () => {
    const node = requireTable(
      "<w:tbl>" + grid(1000) + row(cell("", "a")) + "</w:tbl>"
    );
    expect(node.child(0).attrs.tblPrEx).toBeNull();
  });
});

describe("a cell wrapped in a content control", () => {
  const sdt = (inner: string, pr = '<w:sdtPr><w:id w:val="7"/></w:sdtPr>') =>
    `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

  it("unwraps the cell and holds on to the wrapper", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        `<w:tr>${sdt(cell("", "value"))}${cell("", "a")}</w:tr>` +
        "</w:tbl>"
    );
    expect(texts(node)).toEqual(["value|a"]);
    expect(spans(node)).toEqual(["1x1 1x1"]);
    expect(node.child(0).child(0).attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
    // The cell next to it was never wrapped
    expect(node.child(0).child(1).attrs.sdtPrefix).toBeNull();
  });

  it("the wrapper's own attributes and its sdtEndPr come along", () => {
    const prefix =
      '<w:sdt id="7"><w:sdtPr><w:id w:val="7"/></w:sdtPr>' +
      "<w:sdtEndPr><w:rPr/></w:sdtEndPr>";
    const node = requireTable(
      "<w:tbl>" +
        grid(1000) +
        `<w:tr>${prefix}<w:sdtContent>${cell("", "value")}` +
        "</w:sdtContent></w:sdt></w:tr></w:tbl>"
    );
    expect(node.child(0).child(0).attrs.sdtPrefix).toBe(prefix);
  });

  it("a wrapper that shuts its contents leaves the cell locked", () => {
    const locked =
      '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        `<w:tr>${sdt(cell("", "value"), locked)}${sdt(cell("", "a"))}</w:tr>` +
        "</w:tbl>"
    );
    expect(node.child(0).child(0).attrs.sdtContentsLocked).toBe(true);
    // A control that says nothing about its contents leaves the cell editable
    expect(node.child(0).child(1).attrs.sdtContentsLocked).toBe(false);
  });

  it("a merge count inside the wrapper is read the same way", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        `<w:tr>${sdt(cell('<w:gridSpan w:val="2"/>', "wide"))}</w:tr>` +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["2x1"]);
  });

  /**
   * The cell that starts a vertical merge is a cell of the model in its own right, and the export
   * puts its wrapper back around it, so the table need not be preserved
   */
  it("a cell that starts a vertical merge keeps its wrapper", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        `<w:tr>${sdt(cell('<w:vMerge w:val="restart"/>', "top"))}${cell("", "a")}</w:tr>` +
        row(cell("<w:vMerge/>"), cell("", "b")) +
        "</w:tbl>"
    );
    expect(spans(node)).toEqual(["1x2 1x1", "1x1"]);
    expect(node.child(0).child(0).attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
  });
});

describe("a content control we could not write back out is left preserved", () => {
  const wrapping = (inner: string, ...cols: number[]) =>
    `<w:tbl>${grid(...(cols.length > 0 ? cols : [1000]))}<w:tr>${inner}</w:tr></w:tbl>`;

  it.each([
    [
      "an sdt with no sdtPr",
      wrapping(
        `<w:sdt><w:sdtContent>${cell("", "value")}</w:sdtContent></w:sdt>`
      ),
    ],
    [
      "an sdtContent holding two cells",
      wrapping(
        "<w:sdt><w:sdtPr/><w:sdtContent>" +
          cell("", "a") +
          cell("", "b") +
          "</w:sdtContent></w:sdt>",
        1000,
        1000
      ),
    ],
    [
      "an sdtContent carrying an attribute of its own",
      wrapping(
        `<w:sdt><w:sdtPr/><w:sdtContent w:val="1">${cell("", "value")}</w:sdtContent></w:sdt>`
      ),
    ],
    [
      "an sdt with another element behind its sdtContent",
      wrapping(
        `<w:sdt><w:sdtPr/><w:sdtContent>${cell("", "value")}</w:sdtContent><w:sdtEndPr/></w:sdt>`
      ),
    ],
    // Such a cell is not modelled at all but created fresh on export, wrapper and all
    [
      "a wrapped cell that continues a vertical merge",
      "<w:tbl>" +
        grid(1000) +
        row(cell('<w:vMerge w:val="restart"/>', "top")) +
        "<w:tr><w:sdt><w:sdtPr/><w:sdtContent>" +
        cell("<w:vMerge/>") +
        "</w:sdtContent></w:sdt></w:tr>" +
        "</w:tbl>",
    ],
  ])("%s", (_name, xml) => {
    expect(table(xml)).toBeNull();
  });
});

describe("a table that cannot be modelled is left preserved", () => {
  it.each([
    [
      "a vMerge with no cell to continue from",
      "<w:tbl>" + grid(1000) + row(cell("<w:vMerge/>")) + "</w:tbl>",
    ],
    [
      "a vMerge whose merge width does not line up",
      "<w:tbl>" +
        grid(1000, 1000) +
        row(cell('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>', "top")) +
        row(cell("<w:vMerge/>"), cell("", "a")) +
        "</w:tbl>",
    ],
    [
      "a table whose grid and cell counts do not match",
      "<w:tbl>" +
        grid(1000, 1000, 1000) +
        row(cell("", "a"), cell("", "b")) +
        "</w:tbl>",
    ],
    [
      "a table whose rows have different cell counts",
      "<w:tbl>" +
        row(cell("", "a"), cell("", "b")) +
        row(cell("", "c")) +
        "</w:tbl>",
    ],
    [
      "a row with not a single cell to build",
      "<w:tbl>" +
        grid(1000) +
        row(cell('<w:vMerge w:val="restart"/>', "top")) +
        row("<w:tc><w:tcPr><w:vMerge/></w:tcPr></w:tc>") +
        "</w:tbl>",
    ],
    [
      "a table with an unknown element among the table's children",
      "<w:tbl>" +
        grid(1000) +
        "<w:sdt><w:sdtContent/></w:sdt>" +
        row(cell("", "a")) +
        "</w:tbl>",
    ],
    [
      "a table with an unknown element inside a row",
      "<w:tbl>" +
        grid(1000) +
        "<w:tr><w:customXml/>" +
        cell("", "a") +
        "</w:tr>" +
        "</w:tbl>",
    ],
    ["a table with no rows", "<w:tbl>" + grid(1000) + "</w:tbl>"],
  ])("%s", (_name, xml) => {
    expect(table(xml)).toBeNull();
  });
});

describe("a block inside a cell that cannot be modelled", () => {
  it("a nested table holds on to its original XML as it is", () => {
    const nested =
      '<w:tbl><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid>' +
      "<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
    const node = requireTable(
      "<w:tbl>" +
        grid(1000) +
        `<w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>${nested}</w:tc></w:tr>` +
        "</w:tbl>"
    );
    const cellNode = node.child(0).child(0);
    expect(cellNode.childCount).toBe(2);
    expect(cellNode.child(1).type.name).toBe("rawBlock");
    expect(cellNode.child(1).attrs.xml).toBe(nested);
    expect(cellNode.child(1).attrs.name).toBe("w:tbl");
  });
});

describe("table width and grid", () => {
  it("reads tblW and tcW as a structure", () => {
    const node = requireTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblW w:type="dxa" w:w="6500"/></w:tblPr>' +
        grid(1000, 5500) +
        row(cell('<w:tcW w:type="dxa" w:w="1000"/>', "a"), cell("", "b")) +
        "</w:tbl>"
    );
    expect(node.attrs.tblW).toEqual({ type: "dxa", twips: 6500 });
    expect(node.attrs.gridCols).toEqual([1000, 5500]);
    expect(node.child(0).child(0).attrs.tcW).toEqual({
      type: "dxa",
      twips: 1000,
    });
    // A cell with no tcW is null
    expect(node.child(0).child(1).attrs.tcW).toBeNull();
  });

  it("normalizes percentage widths to a single unit of a fiftieth of a percent", () => {
    const pct = (value: string) =>
      requireTable(
        "<w:tbl>" +
          `<w:tblPr><w:tblW w:type="pct" w:w="${value}"/></w:tblPr>` +
          grid(100) +
          row(cell("", "a")) +
          "</w:tbl>"
      ).attrs.tblW;
    expect(pct("100%")).toEqual({ type: "pct", fiftieths: 5000 });
    expect(pct("5000")).toEqual({ type: "pct", fiftieths: 5000 });
    expect(pct("50%")).toEqual({ type: "pct", fiftieths: 2500 });
  });

  it("reads the widths that leave the size to the layout, number and all", () => {
    const width = (attrs: string) =>
      requireTable(
        "<w:tbl>" +
          `<w:tblPr><w:tblW ${attrs}/></w:tblPr>` +
          grid(1000) +
          row(cell("", "a")) +
          "</w:tbl>"
      ).attrs.tblW;
    // Whatever number stands beside `auto` or `nil` says nothing, so it is not carried around
    expect(width('w:type="auto" w:w="0"')).toEqual({ type: "auto" });
    expect(width('w:type="auto" w:w="5000"')).toEqual({ type: "auto" });
    expect(width('w:type="nil" w:w="0"')).toEqual({ type: "nil" });
  });

  it("an unrecognized width is null, so the original goes back out unchanged", () => {
    const node = requireTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblW w:type="oddValue" w:w="10"/></w:tblPr>' +
        grid(100) +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    expect(node.attrs.tblW).toBeNull();
    expect(node.attrs.tblPr).toContain("oddValue");
  });

  it("does not count revision markup in tblGrid as a column width", () => {
    const node = requireTable(
      "<w:tbl>" +
        '<w:tblGrid><w:gridCol w:w="1000"/>' +
        '<w:tblGridChange w:id="0"><w:tblGrid><w:gridCol w:w="900"/>' +
        "</w:tblGrid></w:tblGridChange></w:tblGrid>" +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    expect(node.attrs.gridCols).toEqual([1000]);
  });
});

describe("table formatting", () => {
  const bordered =
    "<w:tblBorders>" +
    '<w:top w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="4" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="8" w:color="FF0000"/>' +
    '<w:insideV w:val="single" w:sz="8" w:color="00FF00"/>' +
    "</w:tblBorders>";

  it("reads the table borders, alignment and indent", () => {
    const node = requireTable(
      "<w:tbl>" +
        `<w:tblPr>${bordered}<w:jc w:val="center"/>` +
        '<w:tblInd w:type="dxa" w:w="720"/></w:tblPr>' +
        grid(1000) +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    expect(node.attrs.format).toEqual({
      borderTop: "0.5pt solid #000000",
      borderBottom: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      borderRight: "0.5pt solid #000000",
      align: "center",
      indentLeftPt: 36,
    });
  });

  it("reads an indent of nothing, and one that pushes the table off the margin", () => {
    const indent = (twips: string) =>
      requireTable(
        "<w:tbl>" +
          `<w:tblPr><w:tblInd w:type="dxa" w:w="${twips}"/></w:tblPr>` +
          grid(1000) +
          row(cell("", "a")) +
          "</w:tbl>"
      ).attrs.format;
    // Standing at the margin is something the table said, not something it left unsaid
    expect(indent("0")).toEqual({ indentLeftPt: 0 });
    expect(indent("-108")).toEqual({ indentLeftPt: -5.4 });
    expect(
      requireTable("<w:tbl>" + grid(1000) + row(cell("", "a")) + "</w:tbl>")
        .attrs.format
    ).toBeNull();
  });

  it("draws the default thickness for a line that names no thickness of its own", () => {
    const node = requireTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblBorders><w:top w:val="single" w:color="FF0000"/>' +
        '<w:bottom w:val="single" w:sz="0"/></w:tblBorders></w:tblPr>' +
        grid(1000) +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    // Switching such a line off would take away a line the style laid down
    expect(node.attrs.format).toEqual({
      borderTop: "0.5pt solid #FF0000",
      borderBottom: "0.5pt solid #000000",
    });
  });

  it("tells a shading that fills with nothing from one that says nothing at all", () => {
    const shading = (tblPr: string) =>
      styledTable(
        "<w:tbl>" +
          `<w:tblPr><w:tblStyle w:val="Filled"/>${tblPr}</w:tblPr>` +
          grid(1000) +
          row(cell('<w:shd w:val="clear" w:fill="auto"/>', "a")) +
          "</w:tbl>",
        '<w:style w:type="table" w:styleId="Filled">' +
          '<w:tblPr><w:shd w:val="clear" w:fill="FFFF00"/></w:tblPr></w:style>'
      );
    // The table takes the style's fill off again, which is how Word draws it
    expect(
      shading('<w:shd w:val="clear" w:fill="auto"/>').attrs.format
    ).toEqual({ background: "none" });
    // Saying nothing about a fill leaves the style's fill standing
    expect(shading("").attrs.format).toEqual({ background: "#FFFF00" });
    // A cell reads it the same way
    expect(shading("").child(0).child(0).attrs.format).toEqual({
      background: "none",
    });
  });

  /** A 2x2 table with these tcBorders on its top left cell */
  const quartered = (tcBorders = "") =>
    requireTable(
      "<w:tbl>" +
        `<w:tblPr>${bordered}</w:tblPr>` +
        grid(1000, 1000) +
        row(cell(tcBorders, "a"), cell("", "b")) +
        row(cell("", "c"), cell("", "d")) +
        "</w:tbl>"
    );

  it("a cell draws the table's lines, the outer one where it touches the edge of the grid", () => {
    const node = quartered();
    // The top left cell: the table's border above and to its left, its inside lines facing the rest
    expect(node.child(0).child(0).attrs.format).toEqual({
      borderTop: "0.5pt solid #000000",
      borderBottom: "1pt solid #FF0000",
      borderLeft: "0.5pt solid #000000",
      borderRight: "1pt solid #00FF00",
    });
    expect(node.child(1).child(1).attrs.format).toEqual({
      borderTop: "1pt solid #FF0000",
      borderBottom: "0.5pt solid #000000",
      borderLeft: "1pt solid #00FF00",
      borderRight: "0.5pt solid #000000",
    });
  });

  it("a border the cell wrote down beats the lines the table lays down", () => {
    const node = quartered(
      "<w:tcBorders>" +
        '<w:top w:val="nil"/>' +
        '<w:right w:val="double" w:sz="12" w:color="0000FF"/>' +
        "</w:tcBorders>"
    );
    expect(node.child(0).child(0).attrs.format).toMatchObject({
      // The cell switched off the border around the table
      borderTop: "none",
      borderRight: "1.5pt double #0000FF",
      borderBottom: "1pt solid #FF0000",
    });
  });

  it("a table with lines between its cells only draws nothing around the outside", () => {
    const node = requireTable(
      "<w:tbl>" +
        "<w:tblPr><w:tblBorders>" +
        '<w:insideH w:val="single" w:sz="8" w:color="FF0000"/>' +
        '<w:insideV w:val="single" w:sz="8" w:color="00FF00"/>' +
        "</w:tblBorders></w:tblPr>" +
        grid(1000, 1000) +
        row(cell("", "a"), cell("", "b")) +
        row(cell("", "c"), cell("", "d")) +
        "</w:tbl>"
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      borderBottom: "1pt solid #FF0000",
      borderRight: "1pt solid #00FF00",
    });
  });

  it("a merged cell counts the edge it reaches through the merge", () => {
    const node = requireTable(
      "<w:tbl>" +
        `<w:tblPr>${bordered}</w:tblPr>` +
        grid(1000, 1000) +
        row(cell('<w:vMerge w:val="restart"/>', "a"), cell("", "b")) +
        row(cell("<w:vMerge/>"), cell('<w:gridSpan w:val="1"/>', "d")) +
        "</w:tbl>"
    );
    // The rowspan carries the cell down to the last row, so it draws the table's bottom border
    expect(node.child(0).child(0).attrs.format).toEqual({
      borderTop: "0.5pt solid #000000",
      borderBottom: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      borderRight: "1pt solid #00FF00",
    });
  });

  /** The four sides of the top left cell of a 2x2 table whose first cell carries these tcBorders */
  const cellBorders = (sides: string) =>
    quartered(`<w:tcBorders>${sides}</w:tcBorders>`).child(0).child(0).attrs
      .format;

  it("rounds a line style CSS has no name for to the nearest kind it does", () => {
    expect(
      cellBorders(
        '<w:top w:val="triple" w:sz="12" w:color="0000FF"/>' +
          '<w:bottom w:val="dotDotDash" w:sz="8" w:color="0000FF"/>' +
          '<w:left w:val="threeDEmboss" w:sz="8" w:color="0000FF"/>' +
          '<w:right w:val="dashSmallGap" w:sz="8" w:color="0000FF"/>'
      )
    ).toMatchObject({
      borderTop: "1.5pt double #0000FF",
      borderBottom: "1pt dashed #0000FF",
      borderLeft: "1pt solid #0000FF",
      borderRight: "1pt dashed #0000FF",
    });
  });

  it("falls back to the lines the table lays down for a line style it does not know", () => {
    expect(
      cellBorders(
        '<w:top w:val="apples" w:sz="12" w:color="0000FF"/>' +
          '<w:bottom w:val="madeUpKind" w:sz="12" w:color="0000FF"/>'
      )
    ).toMatchObject({
      borderTop: "0.5pt solid #000000",
      borderBottom: "1pt solid #FF0000",
    });
  });

  it("reads the cell's shading, vertical alignment and padding", () => {
    const node = requireTable(
      "<w:tbl>" +
        grid(1000) +
        row(
          cell(
            '<w:shd w:val="clear" w:fill="bfbfbf"/><w:vAlign w:val="center"/>' +
              "<w:tcMar>" +
              '<w:top w:type="dxa" w:w="60"/><w:left w:type="dxa" w:w="100"/>' +
              '<w:bottom w:type="dxa" w:w="60"/><w:right w:type="dxa" w:w="100"/>' +
              "</w:tcMar>",
            "a"
          )
        ) +
        "</w:tbl>"
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      background: "#bfbfbf",
      verticalAlign: "center",
      paddingTopPt: 3,
      paddingRightPt: 5,
      paddingBottomPt: 3,
      paddingLeftPt: 5,
    });
  });

  it("reads a row height telling a minimum apart from an exact one", () => {
    const height = (trPr: string) =>
      requireTable(
        "<w:tbl>" +
          grid(1000) +
          `<w:tr><w:trPr>${trPr}</w:trPr>${cell("", "a")}</w:tr>` +
          "</w:tbl>"
      ).child(0).attrs.format;
    expect(height('<w:trHeight w:val="540" w:hRule="atLeast"/>')).toEqual({
      height: { rule: "atLeast", pt: 27 },
    });
    expect(height('<w:trHeight w:val="540" w:hRule="exact"/>')).toEqual({
      height: { rule: "exact", pt: 27 },
    });
    // With no hRule it is a minimum height
    expect(height('<w:trHeight w:val="540"/>')).toEqual({
      height: { rule: "atLeast", pt: 27 },
    });
    expect(
      height('<w:cantSplit w:val="0"/><w:tblHeader w:val="off"/>')
    ).toEqual({});
  });

  it("reads rows that repeat as headers or cannot split", () => {
    const format = (trPr: string) =>
      requireTable(
        "<w:tbl>" +
          grid(1000) +
          `<w:tr><w:trPr>${trPr}</w:trPr>${cell("", "a")}</w:tr>` +
          "</w:tbl>"
      ).child(0).attrs.format;

    expect(format("<w:tblHeader/><w:cantSplit/>")).toEqual({
      repeatHeader: true,
      cantSplit: true,
    });
    expect(
      format('<w:tblHeader w:val="true"/><w:cantSplit w:val="1"/>')
    ).toEqual({ repeatHeader: true, cantSplit: true });
  });

  it("holds on to the original formatting XML as it is", () => {
    const node = requireTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblStyle w:val="Table1"/></w:tblPr>' +
        grid(1000) +
        `<w:tr><w:trPr><w:cantSplit w:val="0"/></w:trPr>` +
        cell('<w:vAlign w:val="center"/>', "a") +
        "</w:tr></w:tbl>"
    );
    expect(node.attrs.tblPr).toBe(
      '<w:tblPr><w:tblStyle w:val="Table1"/></w:tblPr>'
    );
    expect(node.child(0).attrs.trPr).toBe(
      '<w:trPr><w:cantSplit w:val="0"/></w:trPr>'
    );
    expect(node.child(0).child(0).attrs.tcPr).toBe(
      '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
    );
  });
});

describe("a table whose lines come from a table style", () => {
  const side = (name: string, eighths: number) =>
    `<w:${name} w:val="single" w:sz="${eighths}" w:space="0" w:color="auto"/>`;

  /** Word's Table Grid: a 0.5pt line on all six sides, based on the default table style */
  const TABLE_GRID =
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' +
    '<w:tblPr><w:tblCellMar><w:left w:type="dxa" w:w="108"/></w:tblCellMar>' +
    "</w:tblPr></w:style>" +
    '<w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="TableNormal"/>' +
    "<w:tblPr><w:tblBorders>" +
    ["top", "left", "bottom", "right", "insideH", "insideV"]
      .map((name) => side(name, 4))
      .join("") +
    "</w:tblBorders></w:tblPr></w:style>";

  const GRID_LINE = "0.5pt solid #000000";

  /** The padding the cell margin of the default style above gives every cell: 108 twips */
  const GRID_PADDING = { paddingLeftPt: 5.4 };

  /** A 2x2 table that writes down nothing but the style it points at */
  const referencing = (tblStyle: string) =>
    "<w:tbl>" +
    `<w:tblPr><w:tblStyle w:val="${tblStyle}"/>` +
    '<w:tblW w:type="auto" w:w="0"/></w:tblPr>' +
    grid(1000, 1000) +
    row(cell("", "a"), cell("", "b")) +
    row(cell("", "c"), cell("", "d")) +
    "</w:tbl>";

  it("draws the grid the style defines, in the cells and around the table", () => {
    const node = styledTable(referencing("TableGrid"), TABLE_GRID);
    expect(node.child(0).child(0).attrs.format).toEqual({
      borderTop: GRID_LINE,
      borderBottom: GRID_LINE,
      borderLeft: GRID_LINE,
      borderRight: GRID_LINE,
      ...GRID_PADDING,
    });
    expect(node.attrs.format).toEqual({
      borderTop: GRID_LINE,
      borderBottom: GRID_LINE,
      borderLeft: GRID_LINE,
      borderRight: GRID_LINE,
    });
    // The inside lines are carried along, so an edit can derive the cell values again
    expect(node.attrs.styleInside).toEqual({
      horizontal: GRID_LINE,
      vertical: GRID_LINE,
    });
    // The cell margins are carried along for the same reason
    expect(node.attrs.styleCellMargins).toEqual({
      topPt: null,
      rightPt: null,
      bottomPt: null,
      leftPt: 5.4,
    });
  });

  it("the original formatting XML is left exactly as it was", () => {
    const node = styledTable(referencing("TableGrid"), TABLE_GRID);
    expect(node.attrs.tblPr).toBe(
      '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:type="auto" w:w="0"/></w:tblPr>'
    );
    expect(node.child(0).child(0).attrs.tcPr).toBeNull();
  });

  it("what the table wrote down itself beats the style", () => {
    const node = styledTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblBorders>' +
        '<w:top w:val="single" w:sz="24" w:color="FF0000"/>' +
        '<w:insideH w:val="none"/>' +
        "</w:tblBorders></w:tblPr>" +
        grid(1000) +
        row(cell("", "a")) +
        row(cell("", "b")) +
        "</w:tbl>",
      TABLE_GRID
    );
    expect(node.attrs.format).toMatchObject({
      borderTop: "3pt solid #FF0000",
      // A side the table says nothing about still comes from the style
      borderBottom: GRID_LINE,
    });
    expect(node.child(0).child(0).attrs.format).toMatchObject({
      borderTop: "3pt solid #FF0000",
      // The line between the rows is the one the table switched off
      borderBottom: "none",
      borderLeft: GRID_LINE,
    });
  });

  it("falls back on the default table style when the style it points at is not there", () => {
    const node = styledTable(
      referencing("Missing"),
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' +
        '<w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="999999"/>' +
        "</w:tblBorders></w:tblPr></w:style>",
      "TableNormal"
    );
    expect(node.attrs.styleInside).toEqual({
      horizontal: "0.5pt solid #999999",
      vertical: null,
    });
    // The style draws between the rows and nowhere else, so nothing is drawn around the table
    expect(node.child(0).child(0).attrs.format).toEqual({
      borderBottom: "0.5pt solid #999999",
    });
    expect(node.child(1).child(0).attrs.format).toEqual({
      borderTop: "0.5pt solid #999999",
    });
  });

  it("a table pointing at no style at all takes the default style too", () => {
    const node = styledTable(
      "<w:tbl>" + grid(1000) + row(cell("", "a")) + "</w:tbl>",
      TABLE_GRID,
      "TableGrid"
    );
    expect(node.attrs.styleInside).toEqual({
      horizontal: GRID_LINE,
      vertical: GRID_LINE,
    });
  });

  it("with no styles to look in, only what the table wrote down is drawn", () => {
    const node = requireTable(referencing("TableGrid"));
    // The table has a tblPr, so there is an empty display value, but not a single line in it
    expect(node.attrs.format).toEqual({});
    expect(node.attrs.styleInside).toBeNull();
    expect(node.child(0).child(0).attrs.format).toBeNull();
  });
});

describe("the cell margins a table lays down", () => {
  /** One side of a cell margin, in twips */
  const margin = (name: string, twips: number) =>
    `<w:${name} w:w="${twips}" w:type="dxa"/>`;

  const cellMar = (...sides: string[]) =>
    `<w:tblCellMar>${sides.join("")}</w:tblCellMar>`;

  /** The margins almost every table carries: none above or below, 108 twips at the sides */
  const WORD_MARGINS = cellMar(
    margin("top", 0),
    margin("left", 108),
    margin("bottom", 0),
    margin("right", 108)
  );

  /** The padding those margins turn into */
  const WORD_PADDING = {
    paddingTopPt: 0,
    paddingRightPt: 5.4,
    paddingBottomPt: 0,
    paddingLeftPt: 5.4,
  };

  /** A single-row table whose tblPr writes down nothing but the given cell margins */
  const withMargins = (tblCellMar: string, ...cells: string[]) =>
    "<w:tbl>" +
    `<w:tblPr>${tblCellMar}</w:tblPr>` +
    grid(1000, 1000) +
    row(...cells) +
    "</w:tbl>";

  it("gives every cell the padding the tblPr writes", () => {
    const node = requireTable(
      withMargins(WORD_MARGINS, cell("", "a"), cell("", "b"))
    );
    expect(node.child(0).child(0).attrs.format).toEqual(WORD_PADDING);
    expect(node.child(0).child(1).attrs.format).toEqual(WORD_PADDING);
  });

  it("reads strict leading and trailing margins for left-to-right display", () => {
    const node = requireTable(
      withMargins(
        cellMar(margin("start", 120), margin("end", 180)),
        cell("", "a"),
        cell("", "b")
      )
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      paddingLeftPt: 6,
      paddingRightPt: 9,
    });
  });

  it("lets the cell's own tcMar win, one side at a time", () => {
    const node = requireTable(
      withMargins(
        WORD_MARGINS,
        cell(`<w:tcMar>${margin("left", 400)}</w:tcMar>`, "a"),
        cell("", "b")
      )
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      ...WORD_PADDING,
      // The one side the cell wrote itself, with the other three still the table's
      paddingLeftPt: 20,
    });
    expect(node.child(0).child(1).attrs.format).toEqual(WORD_PADDING);
  });

  it("sets no padding when neither the table nor the cell writes a margin", () => {
    const node = requireTable(
      "<w:tbl>" + grid(1000) + row(cell("", "a")) + "</w:tbl>"
    );
    // Nothing is written down, so the padding on screen is the stylesheet's
    expect(node.child(0).child(0).attrs.format).toBeNull();
  });

  it("reads a margin of nothing, and leaves a unit it cannot use to the level above", () => {
    const node = styledTable(
      withMargins(
        cellMar(
          '<w:top w:w="500" w:type="nil"/>',
          '<w:left w:w="50" w:type="pct"/>'
        ),
        cell("", "a"),
        cell("", "b")
      ),
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' +
        `<w:tblPr>${WORD_MARGINS}</w:tblPr></w:style>`,
      "TableNormal"
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      ...WORD_PADDING,
      // `nil` is a margin of nothing whatever number stands beside it
      paddingTopPt: 0,
      // A share of the page is a unit we cannot turn into points, so the style's margin stands
      paddingLeftPt: 5.4,
    });
  });

  it("layers the tblPr margins on top of the ones the table style laid down", () => {
    const node = styledTable(
      withMargins(cellMar(margin("left", 400)), cell("", "a"), cell("", "b")),
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' +
        `<w:tblPr>${WORD_MARGINS}</w:tblPr></w:style>`,
      "TableNormal"
    );
    expect(node.child(0).child(0).attrs.format).toEqual({
      ...WORD_PADDING,
      paddingLeftPt: 20,
    });
  });
});

describe("tables in the fixtures", () => {
  it.each(fixtureNames)("%s: no table is left as a preserved block", (name) => {
    const { doc } = importDocx(readFixture(name));
    let tables = 0;
    doc.forEach((block) => {
      expect(block.attrs.name).not.toBe("w:tbl");
      if (block.type.name === "table") tables += 1;
    });
    expect(tables).toBeGreaterThan(0);
  });

  it.each(fixtureNames)(
    "%s: prosemirror-tables finds nothing to repair",
    (name) => {
      const { doc } = importDocx(readFixture(name));
      const state = EditorState.create({ doc, schema: docxSchema });
      expect(fixTables(state)).toBeUndefined();
    }
  );
});
