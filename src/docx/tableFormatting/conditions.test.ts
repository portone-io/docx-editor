// @vitest-environment jsdom
/**
 * Which parts of a table a cell belongs to, and what the conditional formatting of its style
 * therefore lays down for it.
 */

import { describe, expect, it } from "vitest";
import {
  type BandSizes,
  type CellStyleFormat,
  TABLE_STYLE_CONDITIONS,
  type TableStyleConditions,
  type TableStyleOverrideType,
} from "../../model/format";
import { parsePropsXml } from "../../ooxml/props";
import { borderLineOfCss } from "../../ooxml/units";
import {
  cellDefaultsFor,
  cellPlacement,
  NO_CELL_SOURCES,
  NO_LOOK,
  readTableLook,
  type TableCellSources,
  type TableLook,
} from "./conditions";
import { NO_BAND_SIZES, NO_CELL_STYLE_FORMAT } from "./reading";

const look = (attrs: string): TableLook =>
  readTableLook(parsePropsXml(`<w:tblPr><w:tblLook ${attrs}/></w:tblPr>`));

const SINGLE_BANDS: BandSizes = { row: null, col: null };

/** The parts of a 4x4 table the cell at this spot belongs to, lowest first */
const conditionsAt = (
  row: number,
  col: number,
  view: TableLook,
  bands: BandSizes = SINGLE_BANDS,
  size = { rows: 4, cols: 4 }
): TableStyleOverrideType[] =>
  cellPlacement(
    { top: row, bottom: row + 1, left: col, right: col + 1 },
    size,
    view,
    bands
  ).conditions.map((condition) => condition.type);

const ALL_PARTS =
  'w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1"';

describe("which parts of its style a table takes", () => {
  it("reads the six attributes", () => {
    expect(look(`${ALL_PARTS} w:noHBand="1" w:noVBand="0"`)).toEqual({
      firstRow: true,
      lastRow: true,
      firstColumn: true,
      lastColumn: true,
      noHBand: true,
      noVBand: false,
    });
  });

  it("takes the banding and nothing else from a table that writes no tblLook at all", () => {
    expect(readTableLook(parsePropsXml("<w:tblPr/>"))).toEqual(NO_LOOK);
    expect(readTableLook(null)).toEqual(NO_LOOK);
    expect(NO_LOOK.noHBand).toBe(false);
  });

  it("reads the legacy hexadecimal bitmask a Transitional document may write instead", () => {
    // 0x0020 first row, 0x0080 first column, 0x0400 no column banding
    expect(look('w:val="04A0"')).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true,
    });
    // Every bit off is the same as writing no tblLook: the rows and columns are still banded
    expect(look('w:val="0000"')).toEqual(NO_LOOK);
    expect(look('w:val="0660"')).toEqual({
      firstRow: true,
      lastRow: true,
      firstColumn: false,
      lastColumn: false,
      noHBand: true,
      noVBand: true,
    });
  });

  it("reads a bitmask that is no bitmask at all as every bit off", () => {
    expect(look('w:val="04A"')).toEqual(NO_LOOK);
    expect(look('w:val="nonsense"')).toEqual(NO_LOOK);
  });

  it("lets an attribute answer for the bit of the same name", () => {
    // The attributes are what a document written today says; the bitmask stands behind them
    expect(look('w:val="04A0" w:firstRow="0" w:lastRow="1"')).toMatchObject({
      firstRow: false,
      lastRow: true,
      firstColumn: true,
      noVBand: true,
    });
  });
});

describe("the parts of a table one cell belongs to", () => {
  it("stand in the order the specification lays them over one another", () => {
    // The header row and the first column are dressed as such and banded with nothing
    const corner = conditionsAt(0, 0, look(ALL_PARTS));
    expect(corner).toEqual(["wholeTable", "firstRow", "firstCol", "nwCell"]);
    const banded = conditionsAt(1, 1, look(ALL_PARTS));
    expect(banded).toEqual(["wholeTable", "band1Vert", "band1Horz"]);
    for (const taken of [corner, banded, conditionsAt(3, 3, look(ALL_PARTS))]) {
      const ranks = taken.map((type) => TABLE_STYLE_CONDITIONS.indexOf(type));
      expect(ranks).toEqual([...ranks].sort((a, b) => a - b));
    }
  });

  it("are the whole table and its bands when the table takes nothing else", () => {
    expect(conditionsAt(0, 0, NO_LOOK)).toEqual([
      "wholeTable",
      "band1Vert",
      "band1Horz",
    ]);
    expect(conditionsAt(1, 1, NO_LOOK)).toEqual([
      "wholeTable",
      "band2Vert",
      "band2Horz",
    ]);
  });

  it("leave the bands out where the table switched them off", () => {
    expect(conditionsAt(1, 1, look('w:noHBand="1"'))).toEqual([
      "wholeTable",
      "band2Vert",
    ]);
    expect(conditionsAt(1, 1, look('w:noHBand="1" w:noVBand="1"'))).toEqual([
      "wholeTable",
    ]);
  });

  it("band the rows the header row and the last row are not part of", () => {
    const view = look('w:firstRow="1" w:lastRow="1" w:noVBand="1"');
    // The header row is dressed as the header and is banded with nothing
    expect(conditionsAt(0, 1, view)).toEqual(["wholeTable", "firstRow"]);
    // The first row under it starts the banding
    expect(conditionsAt(1, 1, view)).toEqual(["wholeTable", "band1Horz"]);
    expect(conditionsAt(2, 1, view)).toEqual(["wholeTable", "band2Horz"]);
    expect(conditionsAt(3, 1, view)).toEqual(["wholeTable", "lastRow"]);
  });

  it("count a band as as many rows as the style said", () => {
    const view = look('w:firstRow="1" w:noVBand="1"');
    const bands: BandSizes = { row: 2, col: null };
    const size = { rows: 7, cols: 2 };
    const horizontal = (row: number) =>
      conditionsAt(row, 1, view, bands, size).filter((type) =>
        type.startsWith("band")
      );
    // Rows 1 and 2 make up the first band, rows 3 and 4 the second, and so on
    expect(horizontal(1)).toEqual(["band1Horz"]);
    expect(horizontal(2)).toEqual(["band1Horz"]);
    expect(horizontal(3)).toEqual(["band2Horz"]);
    expect(horizontal(4)).toEqual(["band2Horz"]);
    expect(horizontal(5)).toEqual(["band1Horz"]);
  });

  it("count the column bands from the first column the same way", () => {
    const view = look('w:noHBand="1"');
    const bands: BandSizes = { row: null, col: 2 };
    const vertical = (col: number) =>
      conditionsAt(0, col, view, bands, { rows: 2, cols: 6 }).filter((type) =>
        type.startsWith("band")
      );
    expect(vertical(0)).toEqual(["band1Vert"]);
    expect(vertical(1)).toEqual(["band1Vert"]);
    expect(vertical(2)).toEqual(["band2Vert"]);
    expect(vertical(5)).toEqual(["band1Vert"]);
  });

  it("carry a corner only where the table takes both the row and the column that meet there", () => {
    const corners = (view: TableLook) =>
      [
        conditionsAt(0, 0, view),
        conditionsAt(0, 3, view),
        conditionsAt(3, 0, view),
        conditionsAt(3, 3, view),
      ].map((taken) => taken.filter((type) => type.endsWith("Cell")));
    expect(corners(look(ALL_PARTS))).toEqual([
      ["nwCell"],
      ["neCell"],
      ["swCell"],
      ["seCell"],
    ]);
    // With only the header row taken, no cell stands where a row and a column both dressed meet
    expect(corners(look('w:firstRow="1"'))).toEqual([[], [], [], []]);
    expect(corners(look('w:firstRow="1" w:lastColumn="1"'))).toEqual([
      [],
      ["neCell"],
      [],
      [],
    ]);
  });

  it("reach a merged cell through every row and column it covers", () => {
    const merged = cellPlacement(
      { top: 0, bottom: 4, left: 0, right: 1 },
      { rows: 4, cols: 4 },
      look(ALL_PARTS),
      SINGLE_BANDS
    ).conditions.map((condition) => condition.type);
    // It starts in the header row and reaches the last one, and stands in the first column
    expect(merged).toEqual([
      "wholeTable",
      "firstRow",
      "lastRow",
      "firstCol",
      "nwCell",
      "swCell",
    ]);
  });
});

describe("what a cell falls back on", () => {
  const condition = (values: Partial<CellStyleFormat>): CellStyleFormat => ({
    ...NO_CELL_STYLE_FORMAT,
    ...values,
  });

  const sourcesWith = (
    conditions: TableStyleConditions,
    view: TableLook = NO_LOOK
  ): TableCellSources => ({
    ...NO_CELL_SOURCES,
    look: view,
    bands: NO_BAND_SIZES,
    conditions,
    styleId: "ListTable",
  });

  const HEADER = condition({
    background: "#D9E2F3",
    borders: {
      top: null,
      bottom: "1pt solid #000000",
      left: null,
      right: null,
    },
    inside: { horizontal: null, vertical: "none" },
  });

  const defaultsAt = (row: number, col: number, sources: TableCellSources) =>
    cellDefaultsFor(
      { top: row, bottom: row + 1, left: col, right: col + 1 },
      { rows: 3, cols: 3 },
      sources
    );

  it("draws a conditional line on the edge of the part it dresses and its inside line within it", () => {
    const sources = sourcesWith({ firstRow: HEADER }, look('w:firstRow="1"'));
    const header = defaultsAt(0, 1, sources);
    expect(header.background).toBe("#D9E2F3");
    expect(header.borders.bottom).toEqual(borderLineOfCss("1pt solid #000000"));
    // The header row is one row deep, so the sides facing the cells beside it take its inside line
    expect(header.borders.left).toEqual(borderLineOfCss("none"));
    expect(header.borders.right).toEqual(borderLineOfCss("none"));
    // A row the condition does not cover is left alone
    expect(defaultsAt(1, 1, sources).background).toBeNull();
    expect(defaultsAt(1, 1, sources).borders.bottom).toBeNull();
  });

  it("lays the parts over one another in the order they are taken", () => {
    const sources = sourcesWith(
      {
        wholeTable: condition({ background: "#FFFFFF" }),
        band1Horz: condition({ background: "#EDEDED" }),
        firstRow: condition({ background: "#D9E2F3" }),
        firstCol: condition({ background: "#C0C0C0" }),
        nwCell: condition({ background: "#000000" }),
      },
      look('w:firstRow="1" w:firstColumn="1"')
    );
    // A banded cell takes the band over what the whole table laid down
    expect(defaultsAt(1, 1, sources).background).toBe("#EDEDED");
    // The first column beats the header row, and the corner beats both
    expect(defaultsAt(0, 1, sources).background).toBe("#D9E2F3");
    expect(defaultsAt(1, 0, sources).background).toBe("#C0C0C0");
    expect(defaultsAt(0, 0, sources).background).toBe("#000000");
  });

  it("leaves what no part it belongs to speaks about as the table left it", () => {
    const sources: TableCellSources = {
      ...sourcesWith({ firstRow: HEADER }, look('w:firstRow="1"')),
      outer: { borderTop: "0.5pt solid #999999" },
      margins: { topPt: 1, rightPt: 4, bottomPt: 1, leftPt: 4 },
    };
    const header = defaultsAt(0, 0, sources);
    expect(header.borders.top).toEqual(borderLineOfCss("0.5pt solid #999999"));
    expect(header.margins).toEqual({
      topPt: 1,
      rightPt: 4,
      bottomPt: 1,
      leftPt: 4,
    });
  });

  it("hands the paragraphs inside the cell the style and the parts they resolve against", () => {
    const sources = sourcesWith({ firstRow: HEADER }, look('w:firstRow="1"'));
    expect(defaultsAt(0, 0, sources).placement).toEqual({
      tableStyleId: "ListTable",
      conditions: ["wholeTable", "band1Vert", "firstRow"],
    });
    expect(defaultsAt(1, 1, sources).placement.conditions).toEqual([
      "wholeTable",
      "band2Vert",
      "band1Horz",
    ]);
  });
});
