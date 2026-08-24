// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  exportErrorCode,
  fixtureNames,
  makeDocx,
  readFixture,
} from "../__testing__/docx";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { serializeTable } from "./serializeTable";

const grid = (...cols: number[]) =>
  `<w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;

const cell = (tcPr: string, text = "") =>
  `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ""}` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const row = (...cells: string[]) => `<w:tr>${cells.join("")}</w:tr>`;

/** Opens a document holding nothing but one table's XML and returns that table node */
function openTable(xml: string): PMNode {
  const { doc } = importDocx(makeDocx(xml));
  const table = doc.child(0);
  if (table.type.name !== "table")
    throw new Error("the table could not be modelled");
  return table;
}

/** Changes the attributes of only the cells that match. This is the shortest way to imitate an edit */
function editCells(
  table: PMNode,
  match: (cell: PMNode) => boolean,
  attrs: Record<string, unknown>
): PMNode {
  const rows: PMNode[] = [];
  table.forEach((rowNode) => {
    const cells: PMNode[] = [];
    rowNode.forEach((cellNode) => {
      cells.push(
        match(cellNode)
          ? cellNode.type.create(
              { ...cellNode.attrs, ...attrs },
              cellNode.content
            )
          : cellNode
      );
    });
    rows.push(rowNode.copy(Fragment.fromArray(cells)));
  });
  return table.copy(Fragment.fromArray(rows));
}

const hasText = (text: string) => (cellNode: PMNode) =>
  cellNode.textContent === text;

describe("rewriting the cell formatting", () => {
  it("gridSpan disappears too when a horizontal merge shrinks to one cell", () => {
    const table = openTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(
          cell(
            '<w:tcW w:type="dxa" w:w="2000"/><w:gridSpan w:val="2"/>',
            "wide"
          )
        ) +
        "</w:tbl>"
    );
    expect(table.child(0).child(0).attrs.colspan).toBe(2);

    const narrowed = editCells(table, hasText("wide"), { colspan: 1 });
    const xml = serializeTable(narrowed);
    expect(xml).not.toContain("gridSpan");
    expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
  });

  it("a growing horizontal merge puts gridSpan into its prescribed order slot", () => {
    const table = openTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(
          cell(
            '<w:tcW w:type="dxa" w:w="1000"/><w:vAlign w:val="center"/>',
            "a"
          ),
          cell("", "b")
        ) +
        "</w:tbl>"
    );
    const widened = editCells(table, hasText("a"), { colspan: 2 });
    expect(serializeTable(widened)).toContain(
      '<w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:gridSpan w:val="2"/>' +
        '<w:vAlign w:val="center"/></w:tcPr>'
    );
  });

  it("vMerge disappears too when a vertical merge goes away", () => {
    const table = openTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(cell('<w:vMerge w:val="restart"/>', "top"), cell("", "a")) +
        row(cell("<w:vMerge/>"), cell("", "b")) +
        "</w:tbl>"
    );
    const unmerged = editCells(table, hasText("top"), { rowspan: 1 });
    expect(serializeTable(unmerged)).not.toContain("vMerge");
  });

  it("rewrites the cell width from the value the model holds", () => {
    const table = openTable(
      "<w:tbl>" +
        grid(1000) +
        row(cell('<w:tcW w:type="dxa" w:w="1000"/>', "a")) +
        "</w:tbl>"
    );
    expect(serializeTable(table)).toContain('<w:tcW w:w="1000" w:type="dxa"/>');

    const wider = editCells(table, hasText("a"), {
      tcW: { type: "dxa", twips: 1500 },
    });
    const xml = serializeTable(wider);
    expect(xml).toContain('<w:tcW w:w="1500" w:type="dxa"/>');
    expect(xml).not.toContain('<w:tcW w:w="1000"');
  });

  it("leaves a width we could not read untouched", () => {
    const odd = '<w:tcW w:type="oddValue" w:w="10"/>';
    const table = openTable(
      "<w:tbl>" + grid(1000) + row(cell(odd, "a")) + "</w:tbl>"
    );
    expect(table.child(0).child(0).attrs.tcW).toBeNull();
    expect(serializeTable(table)).toContain(odd);
  });

  it("writes no tcPr at all when no formatting is left", () => {
    const table = openTable(
      "<w:tbl>" +
        grid(1000, 1000) +
        row(cell('<w:gridSpan w:val="2"/>', "wide")) +
        "</w:tbl>"
    );
    const narrowed = editCells(table, hasText("wide"), { colspan: 1 });
    expect(serializeTable(narrowed)).not.toContain("tcPr");
  });
});

describe("putting a vertical merge back", () => {
  const merged = openTable(
    "<w:tbl>" +
      grid(1000, 1000) +
      row(
        cell(
          '<w:tcW w:type="dxa" w:w="1000"/><w:vMerge w:val="restart"/>',
          "PartyB"
        ),
        cell("", "Company")
      ) +
      row(
        cell('<w:tcW w:type="dxa" w:w="1000"/><w:vMerge/>'),
        cell("", "Address")
      ) +
      row(
        cell('<w:tcW w:type="dxa" w:w="1000"/><w:vMerge/>'),
        cell("", "Contact")
      ) +
      "</w:tbl>"
  );

  it("rebuilds the empty cells on the continuing rows", () => {
    const xml = serializeTable(merged);
    expect(xml.match(/<w:vMerge\/>/g)).toHaveLength(2);
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain(
      '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>'
    );
  });

  it("the continuation cells do not carry the starting cell's text along", () => {
    const xml = serializeTable(merged);
    expect(xml.match(/PartyB/g)).toHaveLength(1);
  });

  it("does not pass over it quietly when the merge is longer than the row count", () => {
    const tooLong = editCells(merged, hasText("PartyB"), { rowspan: 9 });
    expect(exportErrorCode(() => serializeTable(tooLong))).toBe(
      "invalid-table"
    );
  });
});

describe("the row and cell wrappers we carry along without reading them", () => {
  const TBL_PR_EX =
    '<w:tblPrEx><w:tblBorders><w:top w:val="none"/></w:tblBorders></w:tblPrEx>';
  const TR_PR = '<w:trPr><w:cantSplit w:val="0"/></w:trPr>';
  const SDT_PREFIX = '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>';

  /** A cell whose paragraph is already written the way the export writes it */
  const exportedCell = (text: string) =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

  const original =
    "<w:tbl>" +
    grid(1000, 1000) +
    `<w:tr>${TBL_PR_EX}${TR_PR}` +
    `${SDT_PREFIX}<w:sdtContent>${exportedCell("value")}</w:sdtContent></w:sdt>` +
    exportedCell("a") +
    "</w:tr></w:tbl>";

  it("writes the property exceptions ahead of the row properties", () => {
    expect(serializeTable(openTable(original))).toContain(
      `<w:tr>${TBL_PR_EX}${TR_PR}`
    );
  });

  it("puts the content control back around the cell that sat inside it", () => {
    const xml = serializeTable(openTable(original));
    expect(xml).toContain(
      `${SDT_PREFIX}<w:sdtContent>${exportedCell("value")}</w:sdtContent></w:sdt>`
    );
    // The cell next to it goes out on its own
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
  });

  it("an imported table goes back out as the very same XML", () => {
    expect(serializeTable(openTable(original))).toBe(original);
  });

  it("an edited cell keeps both wrappers", () => {
    const table = openTable(original);
    const edited = editCells(table, hasText("value"), {
      tcW: { type: "dxa", twips: 1200 },
    });
    const xml = serializeTable(edited);
    expect(xml).toContain(TBL_PR_EX);
    expect(xml).toContain(
      `${SDT_PREFIX}<w:sdtContent><w:tc><w:tcPr>` +
        '<w:tcW w:w="1200" w:type="dxa"/></w:tcPr>'
    );
  });

  it("the cells rebuilt on the continuing rows of a merge carry no wrapper", () => {
    const table = openTable(original);
    const firstRow = table.child(0);
    const wrapped = firstRow.child(0);
    const plain = firstRow.child(1);
    // The wrapped cell now reaches down into a second row, where only the plain cell stands
    const spanning = wrapped.type.create(
      { ...wrapped.attrs, rowspan: 2 },
      wrapped.content
    );
    const stretched = table.copy(
      Fragment.fromArray([
        firstRow.copy(Fragment.fromArray([spanning, plain])),
        firstRow.copy(Fragment.fromArray([plain])),
      ])
    );
    const xml = serializeTable(stretched);
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
    expect(xml).toContain("<w:vMerge/></w:tcPr><w:p/></w:tc>");
  });
});

describe("a cell that starts a vertical merge inside a content control", () => {
  const SDT_PREFIX = '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr>';
  const START_CELL =
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>' +
    '<w:p><w:r><w:t xml:space="preserve">top</w:t></w:r></w:p></w:tc>';
  /** The empty cell the export rebuilds on the continuing row */
  const CONTINUE_CELL =
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>';

  /** A cell of the second column, written the way the export writes it */
  const plainCell = (text: string) =>
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>' +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

  const original =
    "<w:tbl>" +
    grid(1000, 1000) +
    `<w:tr>${SDT_PREFIX}<w:sdtContent>${START_CELL}</w:sdtContent></w:sdt>` +
    `${plainCell("a")}</w:tr>` +
    `<w:tr>${CONTINUE_CELL}${plainCell("b")}</w:tr>` +
    "</w:tbl>";

  it("opens as one cell reaching down two rows and goes back out as the same XML", () => {
    const table = openTable(original);
    expect(table.childCount).toBe(2);
    expect(table.child(0).child(0).attrs.rowspan).toBe(2);
    expect(serializeTable(table)).toBe(original);
  });

  it("wraps the starting cell alone, never the ones rebuilt below it", () => {
    const xml = serializeTable(openTable(original));
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
    expect(xml).toContain(`<w:tr>${CONTINUE_CELL}`);
  });
});

describe("rewriting the table width and the grid", () => {
  it("rebuilds the grid from the column widths and throws the revision markup away", () => {
    const table = openTable(
      "<w:tbl>" +
        '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="500"/>' +
        '<w:tblGridChange w:id="0"><w:tblGrid><w:gridCol w:w="900"/>' +
        "</w:tblGrid></w:tblGridChange></w:tblGrid>" +
        row(cell("", "a"), cell("", "b")) +
        "</w:tbl>"
    );
    const xml = serializeTable(table);
    expect(xml).toContain(
      '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="500"/></w:tblGrid>'
    );
    expect(xml).not.toContain("tblGridChange");
  });

  it("rewrites tblW inside tblPr when the table width changes", () => {
    const table = openTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblStyle w:val="Table1"/>' +
        '<w:tblW w:type="dxa" w:w="6500"/>' +
        '<w:jc w:val="left"/></w:tblPr>' +
        grid(6500) +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    const wider = table.type.create(
      { ...table.attrs, tblW: { type: "dxa", twips: 7000 } },
      table.content
    );
    expect(serializeTable(wider)).toContain(
      '<w:tblPr><w:tblStyle w:val="Table1"/><w:tblW w:w="7000" w:type="dxa"/>' +
        '<w:jc w:val="left"/></w:tblPr>'
    );
  });

  it("a width appearing on a table that had no tblW goes into its order slot", () => {
    const table = openTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblStyle w:val="Table1"/><w:jc w:val="left"/></w:tblPr>' +
        grid(1000) +
        row(cell("", "a")) +
        "</w:tbl>"
    );
    const sized = table.type.create(
      { ...table.attrs, tblW: { type: "dxa", twips: 1000 } },
      table.content
    );
    expect(serializeTable(sized)).toContain(
      '<w:tblStyle w:val="Table1"/><w:tblW w:w="1000" w:type="dxa"/>' +
        '<w:jc w:val="left"/>'
    );
  });

  it("writes a width that leaves the size to the layout with the 0 Word writes", () => {
    const table = openTable(
      "<w:tbl>" +
        '<w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr>' +
        grid(1000) +
        row(cell('<w:tcW w:type="auto" w:w="0"/>', "a")) +
        "</w:tbl>"
    );
    expect(table.attrs.tblW).toEqual({ type: "auto" });
    const xml = serializeTable(table);
    expect(xml).toContain('<w:tblW w:w="0" w:type="auto"/>');
    expect(xml).toContain('<w:tcW w:w="0" w:type="auto"/>');
  });

  it("creates no grid for a table that had none", () => {
    const table = openTable("<w:tbl>" + row(cell("", "a")) + "</w:tbl>");
    expect(serializeTable(table)).not.toContain("tblGrid");
  });
});

/**
 * Reads the table back from the fixture it came out of, with nothing in the body but the table.
 *
 * The whole package is kept, because the values a table style laid down (the lines between cells,
 * the cell margins) live in styles.xml rather than in the table's own XML. Reopening without them
 * would drop those display values, and the comparison would no longer be like for like.
 */
function reopenIn(fixture: string, xml: string): PMNode {
  const parts = unzipSync(readFixture(fixture));
  const document = decode(parts["word/document.xml"]);
  parts["word/document.xml"] = new TextEncoder().encode(
    document.replace(
      /<w:body>[\s\S]*<\/w:body>/,
      () => `<w:body>${xml}</w:body>`
    )
  );
  const { doc } = importDocx(zipSync(parts));
  const table = doc.child(0);
  if (table.type.name !== "table")
    throw new Error("the table could not be modelled");
  return table;
}

function shapeOf(table: PMNode): unknown {
  const rows: unknown[] = [];
  table.forEach((rowNode) => {
    const cells: unknown[] = [];
    rowNode.forEach((cellNode) =>
      cells.push({
        colspan: cellNode.attrs.colspan,
        rowspan: cellNode.attrs.rowspan,
        tcW: cellNode.attrs.tcW,
        format: cellNode.attrs.format,
        text: cellNode.textContent,
      })
    );
    rows.push({ format: rowNode.attrs.format, cells });
  });
  return {
    gridCols: table.attrs.gridCols,
    tblW: table.attrs.tblW,
    format: table.attrs.format,
    rows,
  };
}

describe("reopening a regenerated table gives the same model", () => {
  it.each(fixtureNames)("%s", (name) => {
    const { doc } = importDocx(readFixture(name));
    const tables: PMNode[] = [];
    doc.forEach((block) => {
      if (block.type.name === "table") tables.push(block);
    });
    expect(tables.length).toBeGreaterThan(0);

    for (const table of tables) {
      const xml = serializeTable(table);
      const again = reopenIn(name, xml);
      expect(shapeOf(again)).toEqual(shapeOf(table));
      // Building it once more gives the same text down to the character (regeneration reaches a fixed point)
      expect(serializeTable(again)).toBe(xml);
    }
  });
});

/** The position just after the first character in the first table, and which body block that table is */
function firstCellSpot(doc: PMNode): { index: number; pos: number } {
  let index = -1;
  let offset = -1;
  doc.forEach((block, blockOffset, i) => {
    if (index === -1 && block.type.name === "table") {
      index = i;
      offset = blockOffset;
    }
  });
  if (index === -1) throw new Error("no table");

  let pos = -1;
  doc.child(index).descendants((node, at) => {
    if (pos === -1 && node.isText) pos = offset + 1 + at + 1;
    return pos === -1;
  });
  if (pos === -1) throw new Error("no text inside the table");
  return { index, pos };
}

describe("locality of an edit inside a table cell", () => {
  it("kitchen sink: only that table block is regenerated and the rest is the original bytes", () => {
    const bytes = readFixture("kitchen-sink.docx");
    const { doc, session } = importDocx(bytes);

    // There has to be a body paragraph ahead of the table for locality to be observable
    const spot = firstCellSpot(doc);
    expect(spot.index).toBeGreaterThan(0);

    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("edit", spot.pos));
    const out = exportDocx(edited.doc, session);
    const exported = unzipSync(out);
    const documentXml = decode(exported[session.mainPartPath]);

    const head =
      session.documentPrefix +
      session.blocks
        .slice(0, spot.index)
        .map((block) => block.xml)
        .join("");
    const tail =
      session.blocks
        .slice(spot.index + 1)
        .map((block) => block.xml)
        .join("") + session.documentSuffix;
    expect(documentXml.startsWith(head)).toBe(true);
    expect(documentXml.endsWith(tail)).toBe(true);

    const middle = documentXml.slice(
      head.length,
      documentXml.length - tail.length
    );
    expect(middle).toContain("edit");
    expect(middle.startsWith("<w:tbl>")).toBe(true);

    // Parts outside the body are left untouched
    const original = unzipSync(bytes);
    for (const key of Object.keys(original)) {
      if (key === session.mainPartPath) continue;
      expect(bytesEqual(exported[key], original[key])).toBe(true);
    }

    // Reopening it, the edit is still there inside the table
    const again = importDocx(out);
    const editedTables: PMNode[] = [];
    again.doc.forEach((block) => {
      if (block.type.name === "table") editedTables.push(block);
    });
    expect(editedTables[0]?.textContent).toContain("edit");
  });

  it.each(fixtureNames)(
    "%s: an edit inside a cell does not touch the other tables and paragraphs",
    (name) => {
      const bytes = readFixture(name);
      const { doc, session } = importDocx(bytes);
      const spot = firstCellSpot(doc);
      const state = createEditorState(doc);
      const edited = state.apply(state.tr.insertText("edit", spot.pos));
      const documentXml = decode(
        unzipSync(exportDocx(edited.doc, session))[session.mainPartPath]
      );

      const others = session.blocks.filter(
        (_block, index) => index !== spot.index
      );
      for (const block of others) {
        expect(documentXml).toContain(block.xml);
      }
    }
  );
});

describe("a node that came from outside a table", () => {
  it("does not pass over a block that cannot go inside a table cell quietly", () => {
    const table = openTable(
      "<w:tbl>" + grid(1000) + row(cell("", "a")) + "</w:tbl>"
    );
    const broken = editCells(table, () => true, {}).copy(
      Fragment.fromArray([
        table.child(0).copy(
          Fragment.fromArray([
            table
              .child(0)
              .child(0)
              .copy(
                Fragment.fromArray([
                  docxSchema.nodes.docxRaw.create({ srcId: 0, name: "w:tbl" }),
                ])
              ),
          ])
        ),
      ])
    );
    expect(exportErrorCode(() => serializeTable(broken))).toBe(
      "unsupported-content"
    );
  });
});
