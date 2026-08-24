// @vitest-environment jsdom
/**
 * Changing a cell's borders and shading.
 *
 * What matters here is what survives: the children of the tcPr the job did not name, the sides of
 * the tcBorders it did not name, and the attributes inside a side it had no business deciding.
 * So the tests pin the exact XML text rather than a parsed shape.
 */

import { describe, expect, it } from "vitest";
import {
  ALL_CELL_SIDES,
  type CellFormatEdit,
  cellBorderDefaults,
  drawsOwnCellBorder,
  editCellProps,
  editRowHeight,
  insideBordersOf,
  NO_BORDER_DEFAULTS,
} from "./tableFormatting";

const ALL_BORDERS: CellFormatEdit = {
  kind: "borders",
  line: "single",
  sides: ALL_CELL_SIDES,
};

const NO_BORDERS: CellFormatEdit = {
  kind: "borders",
  line: "none",
  sides: ALL_CELL_SIDES,
};

const tcPr = (children: string) => `<w:tcPr>${children}</w:tcPr>`;

const borders = (sides: string) => `<w:tcBorders>${sides}</w:tcBorders>`;

const SINGLE = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';

/** The XML the edit produced. Throws when the edit was refused, so a test never asserts on null */
function edited(
  current: string | null,
  edit: CellFormatEdit,
  defaults = NO_BORDER_DEFAULTS
): string | null {
  const next = editCellProps(current, edit, defaults);
  if (!next) throw new Error("the edit was refused");
  return next.tcPr;
}

/** The lines a cell in the middle of such a table falls back on, none of its sides being on an edge */
const insideOf = (tblPr: string) =>
  cellBorderDefaults(
    { top: false, bottom: false, left: false, right: false },
    null,
    insideBordersOf(tblPr)
  );

describe("drawing the borders of a cell", () => {
  it("writes the four sides in the order OOXML lays down, for a cell with no formatting yet", () => {
    expect(edited(null, ALL_BORDERS)).toBe(
      tcPr(
        borders(
          `<w:top ${SINGLE}/><w:left ${SINGLE}/>` +
            `<w:bottom ${SINGLE}/><w:right ${SINGLE}/>`
        )
      )
    );
  });

  it("leaves the other children of the tcPr where they were", () => {
    const current = tcPr(
      '<w:tcW w:w="1500" w:type="dxa"/>' +
        '<w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>' +
        '<w:vAlign w:val="center"/>'
    );
    expect(edited(current, ALL_BORDERS)).toBe(
      tcPr(
        '<w:tcW w:w="1500" w:type="dxa"/>' +
          borders(
            `<w:top ${SINGLE}/><w:left ${SINGLE}/>` +
              `<w:bottom ${SINGLE}/><w:right ${SINGLE}/>`
          ) +
          '<w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>' +
          '<w:vAlign w:val="center"/>'
      )
    );
  });

  it("keeps the sides it was not asked about, inner lines included", () => {
    const current = tcPr(
      borders(
        '<w:top w:val="double" w:sz="12" w:color="FF0000"/>' +
          '<w:insideH w:val="single" w:sz="4"/>' +
          '<w:tl2br w:val="single" w:sz="4"/>'
      )
    );
    expect(
      edited(current, { kind: "borders", line: "single", sides: ["left"] })
    ).toBe(
      tcPr(
        borders(
          '<w:top w:val="double" w:sz="12" w:color="FF0000"/>' +
            `<w:left ${SINGLE}/>` +
            '<w:insideH w:val="single" w:sz="4"/>' +
            '<w:tl2br w:val="single" w:sz="4"/>'
        )
      )
    );
  });

  it("keeps the color a side already had, so a preset never repaints a line", () => {
    const current = tcPr(
      borders('<w:top w:val="dashed" w:sz="12" w:color="FF0000"/>')
    );
    expect(
      edited(current, { kind: "borders", line: "single", sides: ["top"] })
    ).toBe(tcPr(borders('<w:top w:val="single" w:sz="4" w:color="FF0000"/>')));
  });

  it("writes into the spelling the document already used for a side", () => {
    const current = tcPr(borders('<w:start w:val="dotted" w:sz="12"/>'));
    expect(
      edited(current, { kind: "borders", line: "single", sides: ["left"] })
    ).toBe(tcPr(borders('<w:start w:val="single" w:sz="4"/>')));
  });

  it("does not leave both spellings of one side behind", () => {
    const current = tcPr(
      borders(
        '<w:left w:val="single" w:sz="4"/><w:start w:val="dotted" w:sz="12"/>'
      )
    );
    expect(
      edited(current, { kind: "borders", line: "single", sides: ["left"] })
    ).toBe(tcPr(borders('<w:left w:val="single" w:sz="4"/>')));
  });
});

describe("changing a row height", () => {
  it("writes a new height as an OOXML at-least floor", () => {
    expect(editRowHeight(null, 24)).toEqual({
      trPr: '<w:trPr><w:trHeight w:val="480" w:hRule="atLeast"/></w:trPr>',
      format: { height: { rule: "atLeast", pt: 24 } },
    });
  });

  it("keeps unrelated row properties and an existing exact rule", () => {
    const current =
      '<w:trPr><w:cantSplit/><w:trHeight w:val="400" w:hRule="exact"/>' +
      "<w:tblHeader/></w:trPr>";
    expect(editRowHeight(current, 30)).toEqual({
      trPr:
        '<w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="exact"/>' +
        "<w:tblHeader/></w:trPr>",
      format: {
        height: { rule: "exact", pt: 30 },
        repeatHeader: true,
        cantSplit: true,
      },
    });
  });

  it("inserts a new height in row-property order", () => {
    expect(
      editRowHeight("<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>", 24)?.trPr
    ).toBe(
      '<w:trPr><w:cantSplit/><w:trHeight w:val="480" w:hRule="atLeast"/>' +
        "<w:tblHeader/></w:trPr>"
    );
  });

  it("refuses invalid and unchanged values", () => {
    expect(editRowHeight(null, 0)).toBeNull();
    expect(editRowHeight(null, Number.NaN)).toBeNull();
    expect(
      editRowHeight(
        '<w:trPr><w:trHeight w:val="480" w:hRule="atLeast"/></w:trPr>',
        24
      )
    ).toBeNull();
  });
});

describe("cell layout formatting", () => {
  it("writes vertical alignment in tcPr order and preserves unrelated properties", () => {
    expect(
      edited(tcPr('<w:tcW w:w="1500" w:type="dxa"/><w:hideMark/>'), {
        kind: "verticalAlign",
        align: "bottom",
      })
    ).toBe(
      tcPr(
        '<w:tcW w:w="1500" w:type="dxa"/><w:vAlign w:val="bottom"/><w:hideMark/>'
      )
    );
  });

  it("writes only the requested margin sides and retains existing side attributes", () => {
    expect(
      edited(
        tcPr('<w:tcMar><w:top w:w="60" w:type="dxa" w:foo="kept"/></w:tcMar>'),
        { kind: "padding", values: { top: 6, right: 8 } }
      )
    ).toBe(
      tcPr(
        '<w:tcMar><w:top w:w="120" w:type="dxa" w:foo="kept"/>' +
          '<w:right w:w="160" w:type="dxa"/></w:tcMar>'
      )
    );
  });

  it("keeps strict leading and trailing margins in schema order", () => {
    expect(
      edited(
        tcPr(
          '<w:tcMar><w:start w:w="80" w:type="dxa"/>' +
            '<w:end w:w="100" w:type="dxa"/></w:tcMar>'
        ),
        { kind: "padding", values: { top: 6, left: 8 } }
      )
    ).toBe(
      tcPr(
        '<w:tcMar><w:top w:w="120" w:type="dxa"/>' +
          '<w:start w:w="160" w:type="dxa"/>' +
          '<w:end w:w="100" w:type="dxa"/></w:tcMar>'
      )
    );
  });

  it("refuses invalid padding without changing the cell", () => {
    expect(
      editCellProps(null, { kind: "padding", values: { left: -1 } })
    ).toBeNull();
  });
});

describe("clearing the borders of a cell", () => {
  it("switches a side off but leaves behind what it will come back with", () => {
    const current = tcPr(
      borders('<w:top w:val="single" w:sz="12" w:color="FF0000"/>')
    );
    expect(edited(current, NO_BORDERS)).toBe(
      tcPr(
        borders(
          '<w:top w:val="none" w:sz="12" w:color="FF0000"/>' +
            '<w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>'
        )
      )
    );
  });

  it("pins the off state down even where the table draws the inner lines", () => {
    const tblPr =
      "<w:tblPr><w:tblBorders>" +
      '<w:insideH w:val="single" w:sz="4" w:color="999999"/>' +
      '<w:insideV w:val="single" w:sz="4" w:color="999999"/>' +
      "</w:tblBorders></w:tblPr>";
    const next = editCellProps(null, NO_BORDERS, insideOf(tblPr));

    // The cell's own `none` wins over the line the table lays down
    expect(next?.format).toMatchObject({
      borderTop: "none",
      borderBottom: "none",
      borderLeft: "none",
      borderRight: "none",
    });
  });

  it("pins it down on the sides that fall on the border around the table too", () => {
    const line = "0.5pt solid #000000";
    const corner = cellBorderDefaults(
      { top: true, bottom: false, left: true, right: false },
      {
        borderTop: line,
        borderBottom: line,
        borderLeft: line,
        borderRight: line,
      },
      { horizontal: "0.5pt solid #999999", vertical: "0.5pt solid #999999" }
    );
    const next = editCellProps(null, NO_BORDERS, corner);

    expect(next?.format).toMatchObject({
      borderTop: "none",
      borderBottom: "none",
      borderLeft: "none",
      borderRight: "none",
    });
  });
});

describe("the lines a cell falls back on", () => {
  const OUTER = {
    borderTop: "1pt solid #000000",
    borderBottom: "1pt solid #000000",
    borderLeft: "1pt solid #000000",
    borderRight: "1pt solid #000000",
  };
  const INSIDE = {
    horizontal: "0.5pt solid #999999",
    vertical: "0.5pt solid #999999",
  };

  it("are the table's own lines, chosen by where the cell sits in the grid", () => {
    expect(
      cellBorderDefaults(
        { top: true, bottom: false, left: true, right: false },
        OUTER,
        INSIDE
      )
    ).toEqual({
      top: "1pt solid #000000",
      bottom: "0.5pt solid #999999",
      left: "1pt solid #000000",
      right: "0.5pt solid #999999",
    });
  });

  it("draw nothing around a table that only has lines between its cells", () => {
    expect(
      cellBorderDefaults(
        { top: true, bottom: true, left: true, right: true },
        {},
        INSIDE
      )
    ).toEqual(NO_BORDER_DEFAULTS);
  });

  it("keep an outer side the table switched off switched off", () => {
    expect(
      cellBorderDefaults(
        { top: true, bottom: false, left: false, right: false },
        { borderTop: "none" },
        INSIDE
      ).top
    ).toBe("none");
  });
});

describe("coloring the borders of a cell", () => {
  const RED: CellFormatEdit = { kind: "borderColor", hex: "#ff0000" };

  it("recolors only the sides that draw a line", () => {
    const current = tcPr(
      borders(
        '<w:top w:val="single" w:sz="4" w:color="000000"/>' +
          '<w:bottom w:val="nil"/>' +
          '<w:left w:val="dotted" w:sz="8"/>'
      )
    );
    expect(edited(current, RED)).toBe(
      tcPr(
        borders(
          '<w:top w:val="single" w:sz="4" w:color="FF0000"/>' +
            '<w:bottom w:val="nil"/>' +
            '<w:left w:val="dotted" w:sz="8" w:color="FF0000"/>'
        )
      )
    );
  });

  it("drops the theme color, which would otherwise override the one just written", () => {
    const current = tcPr(
      borders(
        '<w:top w:val="single" w:sz="4" w:color="000000" w:themeColor="text1" w:themeTint="80"/>'
      )
    );
    expect(edited(current, RED)).toBe(
      tcPr(borders('<w:top w:val="single" w:sz="4" w:color="FF0000"/>'))
    );
  });

  it("resets the color to auto", () => {
    const current = tcPr(
      borders('<w:top w:val="single" w:sz="4" w:color="FF0000"/>')
    );
    expect(edited(current, { kind: "borderColor", hex: null })).toBe(
      tcPr(borders('<w:top w:val="single" w:sz="4" w:color="auto"/>'))
    );
  });

  it("materializes visible inherited sides without changing their width or style", () => {
    expect(
      edited(null, RED, {
        top: "1.5pt double #A6B7C8",
        bottom: "none",
        left: null,
        right: "0.5pt dotted #A6B7C8",
      })
    ).toBe(
      tcPr(
        borders(
          '<w:top w:val="double" w:sz="12" w:space="0" w:color="FF0000"/>' +
            '<w:right w:val="dotted" w:sz="4" w:space="0" w:color="FF0000"/>'
        )
      )
    );
  });

  it("has nothing to do when neither the cell nor its table draws a line", () => {
    expect(editCellProps(null, RED)).toBeNull();
    expect(
      editCellProps(tcPr(borders('<w:top w:val="nil"/>')), RED)
    ).toBeNull();
    expect(drawsOwnCellBorder(null)).toBe(false);
    expect(drawsOwnCellBorder(tcPr(borders('<w:top w:val="nil"/>')))).toBe(
      false
    );
    expect(
      drawsOwnCellBorder(tcPr(borders('<w:start w:val="single" w:sz="4"/>')))
    ).toBe(true);
  });

  it("leaves an inherited border untouched when it already has the requested color", () => {
    expect(
      editCellProps(null, RED, {
        top: "0.5pt solid #FF0000",
        bottom: null,
        left: null,
        right: null,
      })
    ).toBeNull();
  });
});

describe("filling a cell", () => {
  const YELLOW: CellFormatEdit = { kind: "background", hex: "#ffff00" };

  it("writes a plain fill for a cell that had no shading", () => {
    expect(edited(null, YELLOW)).toBe(
      tcPr('<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')
    );
  });

  it("changes only the fill of the shading that was already there", () => {
    const current = tcPr(
      '<w:shd w:val="pct25" w:color="0000FF" w:fill="D9D9D9" w:themeFill="accent1"/>'
    );
    expect(edited(current, YELLOW)).toBe(
      tcPr('<w:shd w:val="pct25" w:color="0000FF" w:fill="FFFF00"/>')
    );
  });

  it("gives a shading that painted nothing something to paint with", () => {
    const current = tcPr('<w:shd w:val="nil"/>');
    expect(edited(current, YELLOW)).toBe(
      tcPr('<w:shd w:val="clear" w:fill="FFFF00"/>')
    );
  });

  it("drops a shading that recorded nothing but a fill", () => {
    const current = tcPr(
      '<w:tcW w:w="1500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
    );
    expect(edited(current, { kind: "background", hex: null })).toBe(
      tcPr('<w:tcW w:w="1500" w:type="dxa"/>')
    );
  });

  it("keeps a pattern the user did not touch, taking only its fill off", () => {
    const current = tcPr(
      '<w:shd w:val="pct25" w:color="0000FF" w:fill="FFFF00"/>'
    );
    expect(edited(current, { kind: "background", hex: null })).toBe(
      tcPr('<w:shd w:val="pct25" w:color="0000FF" w:fill="auto"/>')
    );
  });

  it("leaves a cell whose formatting held nothing else without a tcPr at all", () => {
    const current = tcPr(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
    );
    expect(edited(current, { kind: "background", hex: null })).toBeNull();
  });
});

describe("refusing an edit", () => {
  it("leaves a cell already in the state the job wants untouched", () => {
    const filled = edited(null, { kind: "background", hex: "#FFFF00" });
    expect(
      editCellProps(filled, { kind: "background", hex: "#ffff00" })
    ).toBeNull();
    const drawn = edited(null, ALL_BORDERS);
    expect(editCellProps(drawn, ALL_BORDERS)).toBeNull();
  });

  it("refuses a color that cannot be written into the document", () => {
    expect(editCellProps(null, { kind: "background", hex: "red" })).toBeNull();
    expect(
      editCellProps(null, { kind: "borderColor", hex: "#12345" })
    ).toBeNull();
  });

  it("refuses formatting XML whose shape could not be made out", () => {
    expect(editCellProps("<w:tcPr><w:shd", ALL_BORDERS)).toBeNull();
    expect(editCellProps(tcPr("<w:tcBorders><w:top"), ALL_BORDERS)).toBeNull();
  });
});

describe("the display values of an edited cell", () => {
  it("come out of the fragment we operated on, the same as on import", () => {
    const next = editCellProps(null, ALL_BORDERS);
    expect(next?.format).toMatchObject({
      borderTop: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
    });

    const filled = editCellProps(next?.tcPr ?? null, {
      kind: "background",
      hex: "#FFFF00",
    });
    // The color keeps the case the document writes it in, which is what the import produces too
    expect(filled?.format).toMatchObject({
      background: "#FFFF00",
      borderTop: "0.5pt solid #000000",
    });
  });

  it("still fold in the lines the table draws between its cells", () => {
    const tblPr =
      '<w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="8" w:color="999999"/>' +
      "</w:tblBorders></w:tblPr>";
    const next = editCellProps(
      null,
      { kind: "background", hex: "#FFFF00" },
      insideOf(tblPr)
    );
    expect(next?.format).toMatchObject({
      background: "#FFFF00",
      borderTop: "1pt solid #999999",
      borderBottom: "1pt solid #999999",
    });
  });
});
