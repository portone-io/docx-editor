// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml } from "../ooxml/xml";
import {
  defaultParagraphStyleIdOf,
  defaultTableStyleIdOf,
  layerParagraphFormat,
  layerRunFormat,
  layerTableFormat,
  readDefaultParagraphFormat,
  readDocumentDefaults,
  readParagraphFormat,
  readParagraphStyles,
  readRunFormat,
  readStyles,
  type StyleFormat,
  styleIdOf,
} from "./formatting";
import { readThemeFonts, type ThemeFonts } from "./theme";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Turns a single XML fragment into an element */
function element(xml: string): Element {
  const wrapped = parseXml(`<w:wrap ${W_NS}>${xml}</w:wrap>`);
  const child = wrapped.documentElement.firstElementChild;
  if (!child) throw new Error("no element");
  return child;
}

const runFormat = (xml: string, themeFonts?: ThemeFonts) =>
  readRunFormat(element(xml), themeFonts);
const paragraphFormat = (xml: string) => readParagraphFormat(element(xml));

/** A font scheme naming a Latin and an East Asian typeface in both of its schemes */
const THEME_FONTS: ThemeFonts = readThemeFonts(
  parseXml(
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      "<a:themeElements><a:fontScheme>" +
      '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface="MS Gothic"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="MS Mincho"/></a:minorFont>' +
      "</a:fontScheme></a:themeElements></a:theme>"
  )
);

/** The style table read out of the given `w:style` elements */
const styleTable = (styles: string, themeFonts?: ThemeFonts) =>
  readStyles(parseXml(`<w:styles ${W_NS}>${styles}</w:styles>`), themeFonts);

/** A whole StyleFormat, with everything the style says nothing about left empty */
const styleFormat = (values: Partial<StyleFormat>): StyleFormat => ({
  paragraph: {},
  run: {},
  table: {},
  tableInside: { horizontal: null, vertical: null },
  tableCellMargins: {
    topPt: null,
    rightPt: null,
    bottomPt: null,
    leftPt: null,
  },
  ...values,
});

/** One border side, as Word writes it inside a table style */
const side = (name: string, eighths: number, color = "auto") =>
  `<w:${name} w:val="single" w:sz="${eighths}" w:space="0" w:color="${color}"/>`;

const tableBorders = (...sides: string[]) =>
  `<w:tblBorders>${sides.join("")}</w:tblBorders>`;

describe("readRunFormat", () => {
  it("is an empty value when there is no formatting", () => {
    expect(runFormat("<w:rPr/>")).toEqual({});
    expect(readRunFormat(null)).toBeNull();
  });

  it("converts a half-point size to points", () => {
    expect(runFormat('<w:rPr><w:sz w:val="20"/></w:rPr>')).toEqual({
      fontSizePt: 10,
    });
    expect(runFormat('<w:rPr><w:sz w:val="22"/></w:rPr>')).toEqual({
      fontSizePt: 11,
    });
    expect(runFormat('<w:rPr><w:sz w:val="19"/></w:rPr>')).toEqual({
      fontSizePt: 9.5,
    });
  });

  it("honors the val of an on/off formatting flag", () => {
    expect(runFormat("<w:rPr><w:b/><w:i/></w:rPr>")).toEqual({
      bold: true,
      italic: true,
    });
    expect(
      runFormat('<w:rPr><w:b w:val="0"/><w:strike w:val="0"/></w:rPr>')
    ).toEqual({});
    expect(runFormat('<w:rPr><w:smallCaps w:val="1"/></w:rPr>')).toEqual({
      smallCaps: true,
    });
  });

  it("keeps the underline kind and leaves out none", () => {
    expect(runFormat('<w:rPr><w:u w:val="single"/></w:rPr>')).toEqual({
      underline: "single",
    });
    expect(runFormat('<w:rPr><w:u w:val="none"/></w:rPr>')).toEqual({});
  });

  it("reads the text color and the highlight", () => {
    expect(runFormat('<w:rPr><w:color w:val="FF0000"/></w:rPr>')).toEqual({
      color: "#FF0000",
    });
    expect(runFormat('<w:rPr><w:color w:val="auto"/></w:rPr>')).toEqual({});
    expect(runFormat('<w:rPr><w:highlight w:val="yellow"/></w:rPr>')).toEqual({
      highlight: "yellow",
    });
    expect(runFormat('<w:rPr><w:highlight w:val="none"/></w:rPr>')).toEqual({});
  });

  it("looks only at the fill color of the shading", () => {
    expect(
      runFormat('<w:rPr><w:shd w:fill="bfbfbf" w:val="clear"/></w:rPr>')
    ).toEqual({ background: "#bfbfbf" });
    // A shading filling with nothing says to paint nothing, which is not the same as saying
    // nothing at all: it takes a fill the style laid down back off
    expect(
      runFormat('<w:rPr><w:shd w:fill="auto" w:val="clear"/></w:rPr>')
    ).toEqual({ background: "none" });
    expect(runFormat("<w:rPr/>")).toEqual({});
  });

  it("collects the font names without duplicates", () => {
    expect(
      runFormat(
        '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic"' +
          ' w:eastAsia="Batang" w:hAnsi="Malgun Gothic"/></w:rPr>'
      )
    ).toEqual({ fontFamily: '"Malgun Gothic","Batang"' });
  });

  /** The name is written into a CSS declaration, and the document naming it is untrusted */
  it("drops a font name that would break the CSS declaration it goes into", () => {
    const dropped = [
      "Arial&quot;;color:red;x:&quot;",
      "Arial;color:red",
      // A trailing backslash would escape the quote the name is wrapped in
      "Arial\\",
    ];
    for (const name of dropped) {
      expect(runFormat(`<w:rPr><w:rFonts w:ascii="${name}"/></w:rPr>`)).toEqual(
        {}
      );
    }
    expect(runFormat('<w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>')).toEqual({
      fontFamily: '"Arial"',
    });
  });

  it("a slot pointing at the theme is drawn in the font the theme names", () => {
    expect(
      runFormat(
        '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi"' +
          ' w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/></w:rPr>',
        THEME_FONTS
      )
    ).toEqual({ fontFamily: '"Calibri","MS Mincho"' });
    expect(
      runFormat(
        '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi"' +
          ' w:eastAsiaTheme="majorEastAsia"/></w:rPr>',
        THEME_FONTS
      )
    ).toEqual({ fontFamily: '"Cambria","MS Gothic"' });
  });

  it("a name in the slot wins over the theme beside it", () => {
    expect(
      runFormat(
        '<w:rPr><w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi"' +
          ' w:eastAsiaTheme="minorEastAsia"/></w:rPr>',
        THEME_FONTS
      )
    ).toEqual({ fontFamily: '"Arial","MS Mincho"' });
  });

  it("a theme nobody handed us leaves the run with no font at all", () => {
    // Which is the document default font, decided by the display layer
    expect(
      runFormat(
        '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi"' +
          ' w:eastAsiaTheme="minorEastAsia"/></w:rPr>'
      )
    ).toEqual({});
  });

  it("leaves out baseline for superscript and subscript", () => {
    expect(
      runFormat('<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>')
    ).toEqual({ verticalAlign: "superscript" });
    expect(runFormat('<w:rPr><w:vertAlign w:val="baseline"/></w:rPr>')).toEqual(
      {}
    );
  });

  it("takes the East Asian language over the Latin one", () => {
    // It is the East Asian language that decides which shape of a Han character is drawn
    expect(
      runFormat('<w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr>')
    ).toEqual({ lang: "ja-JP" });
    expect(runFormat('<w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>')).toEqual({
      lang: "zh-CN",
    });
  });

  it("falls back on the Latin language where the run states no other", () => {
    expect(runFormat('<w:rPr><w:lang w:val="en-US"/></w:rPr>')).toEqual({
      lang: "en-US",
    });
    expect(
      runFormat('<w:rPr><w:lang w:val="en-US" w:eastAsia=""/></w:rPr>')
    ).toEqual({ lang: "en-US" });
    expect(runFormat("<w:rPr><w:lang/></w:rPr>")).toEqual({});
  });

  it("formatting we do not know does not slip into the display values", () => {
    expect(runFormat('<w:rPr><w:rtl w:val="0"/><w:b/></w:rPr>')).toEqual({
      bold: true,
    });
  });
});

describe("readParagraphFormat", () => {
  it("moves the alignment over to a CSS value", () => {
    expect(paragraphFormat('<w:pPr><w:jc w:val="both"/></w:pPr>')).toEqual({
      align: "justify",
    });
    expect(paragraphFormat('<w:pPr><w:jc w:val="center"/></w:pPr>')).toEqual({
      align: "center",
    });
    expect(paragraphFormat('<w:pPr><w:jc w:val="right"/></w:pPr>')).toEqual({
      align: "right",
    });
  });

  it("keeps an explicit paragraph direction in the formatting hierarchy", () => {
    expect(paragraphFormat("<w:pPr><w:bidi/></w:pPr>")).toEqual({
      direction: "rtl",
    });
    expect(paragraphFormat('<w:pPr><w:bidi w:val="0"/></w:pPr>')).toEqual({
      direction: "ltr",
    });
  });

  it("converts the indents from twips to points", () => {
    expect(
      paragraphFormat('<w:pPr><w:ind w:left="1440" w:right="720"/></w:pPr>')
    ).toEqual({ indentStartPt: 72, indentEndPt: 36 });
    expect(
      paragraphFormat('<w:pPr><w:ind w:start="600" w:end="300"/></w:pPr>')
    ).toEqual({ indentStartPt: 30, indentEndPt: 15 });
  });

  it("a hanging indent overrides the first line indent", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:ind w:left="1068" w:firstLine="200" w:hanging="360"/></w:pPr>'
      )
    ).toEqual({ indentStartPt: 53.4, textIndentPt: -18 });
    expect(
      paragraphFormat('<w:pPr><w:ind w:firstLine="192"/></w:pPr>')
    ).toEqual({ textIndentPt: 9.6 });
  });

  it("reads the paragraph spacing and the line spacing", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:spacing w:after="200" w:before="240"/></w:pPr>'
      )
    ).toEqual({ spaceAfterPt: 10, spaceBeforePt: 12 });
    // auto means a multiple in 240ths (276 -> 1.15 lines)
    expect(
      paragraphFormat(
        '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>'
      )
    ).toEqual({ lineSpacing: { rule: "auto", lines: 1.15 } });
    // With no lineRule it is treated as auto (the OOXML default)
    expect(paragraphFormat('<w:pPr><w:spacing w:line="480"/></w:pPr>')).toEqual(
      { lineSpacing: { rule: "auto", lines: 2 } }
    );
    expect(
      paragraphFormat(
        '<w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr>'
      )
    ).toEqual({ lineSpacing: { rule: "exact", pt: 18 } });
    expect(
      paragraphFormat(
        '<w:pPr><w:spacing w:line="360" w:lineRule="atLeast"/></w:pPr>'
      )
    ).toEqual({ lineSpacing: { rule: "atLeast", pt: 18 } });
  });

  it("reads the list position. numId 0 is not a list", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="6"/>' +
          "</w:numPr></w:pPr>"
      )
    ).toEqual({ numbering: { numId: 6, ilvl: 2 } });
    // With no ilvl it is the first level
    expect(
      paragraphFormat('<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>')
    ).toEqual({ numbering: { numId: 3, ilvl: 0 } });
    expect(
      paragraphFormat(
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/>' +
          "</w:numPr></w:pPr>"
      )
    ).toEqual({ numbering: null });
  });

  it("keeps only the border sides that are actually drawn", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:pBdr><w:top w:space="0" w:sz="0" w:val="nil"/>' +
          '<w:bottom w:space="0" w:sz="8" w:val="single" w:color="000000"/>' +
          "</w:pBdr></w:pPr>"
      )
    ).toEqual({ borderBottom: "1pt solid #000000" });
  });

  it("reads the paragraph shading", () => {
    expect(
      paragraphFormat('<w:pPr><w:shd w:fill="bfbfbf" w:val="clear"/></w:pPr>')
    ).toEqual({ background: "#bfbfbf" });
  });

  it("only a paragraph that asks a page to start here carries it in the display values", () => {
    expect(paragraphFormat("<w:pPr><w:pageBreakBefore/></w:pPr>")).toEqual({
      pageBreakBefore: true,
    });
    expect(
      paragraphFormat('<w:pPr><w:pageBreakBefore w:val="1"/></w:pPr>')
    ).toEqual({ pageBreakBefore: true });
    expect(
      paragraphFormat('<w:pPr><w:pageBreakBefore w:val="0"/></w:pPr>')
    ).toEqual({});
  });

  it("formatting we do not handle yet does not slip into the display values", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:keepNext/><w:widowControl w:val="0"/>' +
          '<w:pStyle w:val="Heading1"/></w:pPr>'
      )
    ).toEqual({});
    expect(readParagraphFormat(null)).toBeNull();
  });

  it("reads custom tab directives without resolving their hierarchy", () => {
    expect(
      paragraphFormat(
        '<w:pPr><w:tabs><w:tab w:val="start" w:pos="720" w:leader="dot"/>' +
          '<w:tab w:val="clear" w:pos="1440"/></w:tabs></w:pPr>'
      )
    ).toEqual({
      tabStops: [
        { positionPt: 36, align: "start", leader: "dot" },
        { positionPt: 72, align: "clear" },
      ],
    });
  });
});

describe("readDocumentDefaults", () => {
  const styles = (docDefaults: string, themeFonts?: ThemeFonts) =>
    readDocumentDefaults(
      parseXml(`<w:styles ${W_NS}>${docDefaults}</w:styles>`),
      themeFonts
    );

  it("reads the default size and font out of docDefaults", () => {
    expect(
      styles(
        "<w:docDefaults><w:rPrDefault><w:rPr>" +
          '<w:rFonts w:ascii="Arial" w:eastAsia="Arial"/><w:sz w:val="22"/>' +
          "</w:rPr></w:rPrDefault></w:docDefaults>"
      )
    ).toEqual({ fontSizePt: 11, fontFamily: '"Arial"', lineSpacing: null });
  });

  it("reads the default line spacing out of docDefaults", () => {
    expect(
      styles(
        "<w:docDefaults><w:pPrDefault><w:pPr>" +
          '<w:spacing w:line="276" w:lineRule="auto"/>' +
          "</w:pPr></w:pPrDefault></w:docDefaults>"
      )
    ).toEqual({
      fontSizePt: null,
      fontFamily: null,
      lineSpacing: { rule: "auto", lines: 1.15 },
    });
  });

  it("resolves the default font out of the theme when that is all it names", () => {
    expect(
      styles(
        "<w:docDefaults><w:rPrDefault><w:rPr>" +
          '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"' +
          ' w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/>' +
          "</w:rPr></w:rPrDefault></w:docDefaults>",
        THEME_FONTS
      )
    ).toEqual({
      fontSizePt: 11,
      fontFamily: '"Calibri","MS Mincho"',
      lineSpacing: null,
    });
  });

  it("is null when the defaults are empty", () => {
    const empty = { fontSizePt: null, fontFamily: null, lineSpacing: null };
    expect(
      styles("<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>")
    ).toEqual(empty);
    expect(styles("")).toEqual(empty);
  });
});

describe("default paragraph properties", () => {
  it("reads the complete pPrDefault layer used by the formatting hierarchy", () => {
    const styles = parseXml(
      `<w:styles ${W_NS}><w:docDefaults><w:pPrDefault><w:pPr>` +
        '<w:ind w:left="360"/><w:tabs><w:tab w:val="start" w:pos="720"/>' +
        "</w:tabs></w:pPr></w:pPrDefault></w:docDefaults></w:styles>"
    );

    expect(readDefaultParagraphFormat(styles)).toEqual({
      indentStartPt: 18,
      tabStops: [{ positionPt: 36, align: "start" }],
    });
  });
});

describe("the style chain", () => {
  const style = (id: string, body: string, basedOn?: string) =>
    `<w:style w:styleId="${id}">` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : "") +
    body +
    "</w:style>";

  it("a style pointing at the theme lays down the font the theme names", () => {
    const styles = styleTable(
      style(
        "Normal",
        '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi"' +
          ' w:eastAsiaTheme="minorEastAsia"/></w:rPr>'
      ),
      THEME_FONTS
    );
    expect(styles.get("Normal")?.run).toEqual({
      fontFamily: '"Calibri","MS Mincho"',
    });
  });

  it("reads the paragraph formatting and the run formatting together", () => {
    const styles = styleTable(
      style(
        "Heading1",
        '<w:pPr><w:jc w:val="center"/></w:pPr>' +
          '<w:rPr><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr>'
      )
    );
    expect(styles.get("Heading1")).toEqual(
      styleFormat({
        paragraph: { align: "center" },
        run: { color: "#2E74B5", fontSizePt: 16 },
      })
    );
  });

  it("follows basedOn and layers from the root down, the lower one winning", () => {
    const styles = styleTable(
      style("Normal", '<w:rPr><w:sz w:val="20"/></w:rPr>') +
        style("Body", "<w:rPr><w:b/></w:rPr>", "Normal") +
        style("Note", '<w:rPr><w:sz w:val="18"/></w:rPr>', "Body")
    );
    expect(styles.get("Note")?.run).toEqual({ bold: true, fontSizePt: 9 });
    expect(styles.get("Body")?.run).toEqual({ bold: true, fontSizePt: 10 });
  });

  it("adds, replaces, and clears tab stops through a style chain", () => {
    const tabs = (...entries: string[]) =>
      `<w:pPr><w:tabs>${entries.join("")}</w:tabs></w:pPr>`;
    const tab = (value: string, position: number) =>
      `<w:tab w:val="${value}" w:pos="${position}"/>`;
    const styles = styleTable(
      style("Normal", tabs(tab("start", 720), tab("center", 1440))) +
        style("Body", tabs(tab("end", 1440), tab("decimal", 2160)), "Normal") +
        style("Note", tabs(tab("clear", 720)), "Body")
    );

    expect(styles.get("Note")?.paragraph.tabStops).toEqual([
      { positionPt: 36, align: "clear" },
      { positionPt: 72, align: "end" },
      { positionPt: 108, align: "decimal" },
    ]);
  });

  it("lets direct paragraph properties clear an inherited tab stop", () => {
    expect(
      layerParagraphFormat(
        {
          tabStops: [
            { positionPt: 36, align: "start", leader: "dot" },
            { positionPt: 72, align: "center" },
          ],
        },
        { tabStops: [{ positionPt: 36, align: "clear" }] }
      )
    ).toEqual({ tabStops: [{ positionPt: 72, align: "center" }] });
  });

  it("keeps what it has read so far even when a style points at one that is missing", () => {
    const styles = styleTable(
      style("Heading1", "<w:rPr><w:b/></w:rPr>", "Normal")
    );
    expect(styles.get("Heading1")?.run).toEqual({ bold: true });
    expect(styles.get("Normal")).toBeUndefined();
  });

  it("stops even when basedOn goes round in a circle", () => {
    const styles = styleTable(
      style("A", "<w:rPr><w:b/></w:rPr>", "B") +
        style("B", "<w:rPr><w:i/></w:rPr>", "A")
    );
    expect(styles.get("A")?.run).toEqual({ bold: true, italic: true });
    expect(styles.get("B")?.run).toEqual({ bold: true, italic: true });
  });

  it("leaves out a style that has no styleId", () => {
    expect(styleTable("<w:style><w:rPr><w:b/></w:rPr></w:style>").size).toBe(0);
  });
});

describe("the table formatting a style passes down", () => {
  const tableStyle = (id: string, tblPr: string, basedOn?: string) =>
    `<w:style w:type="table" w:styleId="${id}">` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : "") +
    `<w:tblPr>${tblPr}</w:tblPr>` +
    "</w:style>";

  /** The lines Word's Table Grid draws: 0.5pt on all six sides */
  const gridBorders = tableBorders(
    side("top", 4),
    side("left", 4),
    side("bottom", 4),
    side("right", 4),
    side("insideH", 4),
    side("insideV", 4)
  );

  it("reads the borders out of the style's tblPr", () => {
    const styles = styleTable(tableStyle("TableGrid", gridBorders));
    expect(styles.get("TableGrid")?.table).toEqual({
      borderTop: "0.5pt solid #000000",
      borderBottom: "0.5pt solid #000000",
      borderLeft: "0.5pt solid #000000",
      borderRight: "0.5pt solid #000000",
    });
    expect(styles.get("TableGrid")?.tableInside).toEqual({
      horizontal: "0.5pt solid #000000",
      vertical: "0.5pt solid #000000",
    });
  });

  it("reads the alignment and the shading too", () => {
    const styles = styleTable(
      tableStyle(
        "Centered",
        '<w:jc w:val="center"/><w:shd w:val="clear" w:fill="EEEEEE"/>'
      )
    );
    expect(styles.get("Centered")?.table).toEqual({
      align: "center",
      background: "#EEEEEE",
    });
  });

  it("a style further down the chain covers only the inside lines it writes itself", () => {
    const styles = styleTable(
      tableStyle("TableNormal", gridBorders) +
        tableStyle(
          "Loud",
          tableBorders(side("insideH", 12, "FF0000")),
          "TableNormal"
        )
    );
    expect(styles.get("Loud")?.tableInside).toEqual({
      horizontal: "1.5pt solid #FF0000",
      // The side it says nothing about keeps the line from the style above
      vertical: "0.5pt solid #000000",
    });
  });

  it("a style switches a line from above off by writing that side as none", () => {
    const styles = styleTable(
      tableStyle("TableNormal", gridBorders) +
        tableStyle(
          "NoRules",
          tableBorders('<w:insideH w:val="none"/>', '<w:insideV w:val="nil"/>'),
          "TableNormal"
        )
    );
    expect(styles.get("NoRules")?.tableInside).toEqual({
      horizontal: "none",
      vertical: "none",
    });
  });

  /** One side of a cell margin, in twips, as Word writes it inside a table style */
  const margin = (name: string, twips: number) =>
    `<w:${name} w:w="${twips}" w:type="dxa"/>`;

  const cellMargins = (...sides: string[]) =>
    `<w:tblCellMar>${sides.join("")}</w:tblCellMar>`;

  /** The margins Word's own default table style writes: none above or below, 108 twips at the sides */
  const wordMargins = cellMargins(
    margin("top", 0),
    margin("left", 108),
    margin("bottom", 0),
    margin("right", 108)
  );

  it("reads the cell margins out of the style's tblPr", () => {
    const styles = styleTable(tableStyle("TableNormal", wordMargins));
    expect(styles.get("TableNormal")?.tableCellMargins).toEqual({
      topPt: 0,
      rightPt: 5.4,
      bottomPt: 0,
      leftPt: 5.4,
    });
  });

  it("a style further down the chain covers only the cell margins it writes itself", () => {
    const styles = styleTable(
      tableStyle("TableNormal", wordMargins) +
        tableStyle(
          "Roomy",
          cellMargins(margin("top", 60), margin("left", 200)),
          "TableNormal"
        )
    );
    expect(styles.get("Roomy")?.tableCellMargins).toEqual({
      topPt: 3,
      leftPt: 10,
      // The sides it says nothing about keep the margins from the style above
      rightPt: 5.4,
      bottomPt: 0,
    });
  });

  it("a side written on its own leaves the other three with nothing", () => {
    const styles = styleTable(
      tableStyle("OneSide", cellMargins(margin("left", 108)))
    );
    expect(styles.get("OneSide")?.tableCellMargins).toEqual({
      topPt: null,
      rightPt: null,
      bottomPt: null,
      leftPt: 5.4,
    });
  });

  it("a style with no table formatting has nothing to pass down", () => {
    const styles = styleTable(
      '<w:style w:styleId="Plain"><w:rPr><w:b/></w:rPr></w:style>'
    );
    expect(styles.get("Plain")).toEqual(styleFormat({ run: { bold: true } }));
  });
});

describe("defaultTableStyleIdOf", () => {
  const defaultOf = (styles: string) =>
    defaultTableStyleIdOf(parseXml(`<w:styles ${W_NS}>${styles}</w:styles>`));

  const tableStyle = (id: string, attrs = "") =>
    `<w:style w:type="table" w:styleId="${id}"${attrs}/>`;

  it("finds the table style the document marked as its default", () => {
    expect(
      defaultOf(
        tableStyle("TableGrid") + tableStyle("TableNormal", ' w:default="1"')
      )
    ).toBe("TableNormal");
    expect(defaultOf(tableStyle("TableNormal", ' w:default="true"'))).toBe(
      "TableNormal"
    );
  });

  it("is null when no table style is the default", () => {
    expect(defaultOf(tableStyle("TableGrid"))).toBeNull();
    // A default of another kind is not a default table style
    expect(
      defaultOf(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'
      )
    ).toBeNull();
    expect(defaultOf("")).toBeNull();
  });

  it("the last of several defaults wins, as Word reads it", () => {
    expect(
      defaultOf(
        tableStyle("First", ' w:default="1"') +
          tableStyle("Second", ' w:default="1"')
      )
    ).toBe("Second");
  });
});

describe("defaultParagraphStyleIdOf", () => {
  const defaultOf = (styles: string) =>
    defaultParagraphStyleIdOf(
      parseXml(`<w:styles ${W_NS}>${styles}</w:styles>`)
    );

  const paragraphStyle = (id: string, attrs = "") =>
    `<w:style w:type="paragraph" w:styleId="${id}"${attrs}/>`;

  it("finds the paragraph style the document marked as its default", () => {
    expect(
      defaultOf(
        paragraphStyle("Heading1") + paragraphStyle("Normal", ' w:default="1"')
      )
    ).toBe("Normal");
  });

  it("is null when no paragraph style is the default", () => {
    expect(defaultOf(paragraphStyle("Normal"))).toBeNull();
    // A default of another kind is not a default paragraph style
    expect(
      defaultOf(
        '<w:style w:type="table" w:styleId="TableNormal" w:default="1"/>'
      )
    ).toBeNull();
  });

  it("the last of several defaults wins, the same as for tables", () => {
    expect(
      defaultOf(
        paragraphStyle("First", ' w:default="1"') +
          paragraphStyle("Second", ' w:default="1"')
      )
    ).toBe("Second");
  });
});

describe("readParagraphStyles", () => {
  const optionsOf = (styles: string) =>
    readParagraphStyles(parseXml(`<w:styles ${W_NS}>${styles}</w:styles>`));

  const paragraphStyle = (id: string, body = "", attrs = "") =>
    `<w:style w:type="paragraph" w:styleId="${id}"${attrs}>${body}</w:style>`;

  const name = (value: string) => `<w:name w:val="${value}"/>`;

  it("offers the paragraph styles in the order the document lists them", () => {
    expect(
      optionsOf(
        paragraphStyle("Heading1", name("heading 1")) +
          paragraphStyle("Quote", name("Quote"))
      )
    ).toEqual([
      {
        id: "Heading1",
        name: "heading 1",
        isDefault: false,
        primary: false,
        hidden: false,
      },
      {
        id: "Quote",
        name: "Quote",
        isDefault: false,
        primary: false,
        hidden: false,
      },
    ]);
  });

  it("leaves out the styles of other kinds", () => {
    expect(
      optionsOf(
        '<w:style w:type="table" w:styleId="TableGrid"/>' +
          '<w:style w:type="character" w:styleId="Strong"/>' +
          paragraphStyle("Quote")
      ).map((option) => option.id)
    ).toEqual(["Quote"]);
  });

  it("falls back on the id where the style states no name", () => {
    expect(optionsOf(paragraphStyle("Quote"))).toEqual([
      {
        id: "Quote",
        name: "Quote",
        isDefault: false,
        primary: false,
        hidden: false,
      },
    ]);
    expect(
      optionsOf(paragraphStyle("Quote", '<w:name w:val=""/>'))[0].name
    ).toBe("Quote");
  });

  /**
   * The name is the only thing that can be shown for a style a paragraph points at, so a hidden
   * style is carried along flagged rather than dropped along with its name
   */
  it("flags the styles Word keeps out of its own gallery", () => {
    expect(
      optionsOf(
        paragraphStyle("Index1", `${name("index 1")}<w:semiHidden/>`) +
          paragraphStyle("Hidden1", `${name("hidden 1")}<w:hidden/>`) +
          paragraphStyle("Quote", name("Quote"))
      )
    ).toEqual([
      {
        id: "Index1",
        name: "index 1",
        isDefault: false,
        primary: false,
        hidden: true,
      },
      {
        id: "Hidden1",
        name: "hidden 1",
        isDefault: false,
        primary: false,
        hidden: true,
      },
      {
        id: "Quote",
        name: "Quote",
        isDefault: false,
        primary: false,
        hidden: false,
      },
    ]);
    // A style that pins the flag down as off is offered as any other
    expect(
      optionsOf(paragraphStyle("Quote", '<w:semiHidden w:val="0"/>'))[0].hidden
    ).toBe(false);
  });

  it("marks the primary styles a document declares", () => {
    expect(
      optionsOf(
        paragraphStyle("Heading1", `${name("heading 1")}<w:qFormat/>`) +
          paragraphStyle("Quote", `${name("Quote")}<w:qFormat w:val="0"/>`)
      ).map((option) => option.primary)
    ).toEqual([true, false]);
  });

  it("marks the style a paragraph with no pStyle already wears", () => {
    expect(
      optionsOf(
        paragraphStyle("Normal", name("Normal"), ' w:default="1"') +
          paragraphStyle("Quote", name("Quote"))
      ).map((option) => option.isDefault)
    ).toEqual([true, false]);
    expect(
      optionsOf(paragraphStyle("Normal", "", ' w:default="true"'))[0].isDefault
    ).toBe(true);
    expect(
      optionsOf(paragraphStyle("Normal", "", ' w:default="0"'))[0].isDefault
    ).toBe(false);
  });

  // Word reads the last of several defaults as the one in force, so only that one is marked here
  it("marks only the last of several defaults", () => {
    expect(
      optionsOf(
        paragraphStyle("First", name("First"), ' w:default="1"') +
          paragraphStyle("Second", name("Second"), ' w:default="1"')
      ).map((option) => option.isDefault)
    ).toEqual([false, true]);
  });

  it("leaves out a style with no styleId", () => {
    expect(optionsOf('<w:style w:type="paragraph"/>')).toEqual([]);
    expect(optionsOf("")).toEqual([]);
  });
});

describe("styleIdOf", () => {
  it("pulls out the styleId the paragraph points at", () => {
    expect(
      styleIdOf(
        '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>'
      )
    ).toBe("Heading1");
  });

  it("is null when it points at nothing", () => {
    expect(styleIdOf('<w:pPr><w:jc w:val="center"/></w:pPr>')).toBeNull();
    expect(styleIdOf(null)).toBeNull();
    expect(styleIdOf(42)).toBeNull();
  });

  // The answer is kept per fragment, and "no name in it" is an answer like any other
  it("keeps answering the same for a fragment it has already read", () => {
    const named = '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>';
    expect(styleIdOf(named)).toBe("Quote");
    expect(styleIdOf(named)).toBe("Quote");

    const unreadable = "<w:pPr><w:pStyle/></w:pPr>";
    expect(styleIdOf(unreadable)).toBeNull();
    expect(styleIdOf(unreadable)).toBeNull();
  });
});

describe("layering the style values underneath", () => {
  it("direct formatting covers the style", () => {
    expect(
      layerParagraphFormat(
        { align: "center", spaceAfterPt: 10 },
        { align: "right" }
      )
    ).toEqual({ align: "right", spaceAfterPt: 10 });
    expect(
      layerRunFormat({ color: "#2E74B5", fontSizePt: 16 }, { fontSizePt: 18 })
    ).toEqual({ color: "#2E74B5", fontSizePt: 18 });
    expect(
      layerTableFormat(
        { borderTop: "0.5pt solid #000000", align: "center" },
        { borderTop: "none" }
      )
    ).toEqual({ borderTop: "none", align: "center" });
  });

  it("creates no display values when there is neither a style nor direct formatting", () => {
    expect(layerParagraphFormat({}, null)).toBeNull();
    expect(layerRunFormat({}, null)).toBeNull();
    expect(layerTableFormat({}, null)).toBeNull();
    // When the direct formatting is an empty value, that empty value is kept
    expect(layerRunFormat({}, {})).toEqual({});
  });
});
