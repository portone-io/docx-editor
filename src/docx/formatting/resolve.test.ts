// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeStyledNumberedDocx,
  readFixture,
} from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { toggleNumberedList } from "../../editor/commands/listCommands";
import { setParagraphStyle } from "../../editor/commands/paragraphCommands";
import { editorStateForSession } from "../../editor/createEditor";
import { docxKeymap } from "../../editor/plugins/keymap";
import { toParagraphFormat } from "../../model/format";
import type { Numbering } from "../../numbering/parseNumbering";
import { parseXml } from "../../ooxml/xml";
import { docxSchema } from "../../schema";
import { importDocx } from "../importDocx";
import { paragraphAttrsFor } from "./attrs";
import {
  type FormattingContext,
  formattingContextOf,
  NO_FORMATTING,
} from "./context";
import {
  inheritedRunFormat,
  type ResolvedParagraph,
  resolveParagraph,
  resolveRun,
} from "./resolve";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function context(
  stylesXml: string,
  overrides: Partial<FormattingContext> = {}
): FormattingContext {
  return {
    ...formattingContextOf(
      parseXml(`<w:styles ${W_NS}>${stylesXml}</w:styles>`),
      overrides.numbering ?? NO_FORMATTING.numbering,
      NO_FORMATTING.themeFonts
    ),
    ...overrides,
  };
}

const NUMBERING: Numbering = {
  lists: new Map([
    [
      4,
      {
        levels: new Map([
          [
            2,
            {
              format: "decimal",
              text: "%1.",
              start: 1,
              indent: null,
              tabStops: [
                { positionPt: 36, align: "clear" },
                { positionPt: 54, align: "num" },
              ],
            },
          ],
        ]),
      },
    ],
  ]),
};

const LIST_PPR =
  "<w:numPr>" + '<w:ilvl w:val="2"/><w:numId w:val="4"/></w:numPr>';

describe("the paragraph hierarchy", () => {
  /**
   * Every layer speaks for one thing the layer below it also spoke for, so the order they are
   * laid down in is what the outcome reads back: the document defaults and the table style
   * both space the paragraph, the table style and the numbering level both place a stop at
   * 36pt, the numbering level and the paragraph style both speak for 54pt, and the paragraph
   * style and the direct formatting both align the paragraph.
   */
  it("layers docDefaults, table style, numbering, paragraph style and direct in §17.7.2 order", () => {
    const formatting = context(
      "<w:docDefaults><w:pPrDefault><w:pPr>" +
        '<w:spacing w:before="20" w:after="20"/>' +
        "</w:pPr></w:pPrDefault></w:docDefaults>" +
        '<w:style w:type="table" w:styleId="Grid"><w:pPr>' +
        '<w:spacing w:after="40"/>' +
        '<w:tabs><w:tab w:val="start" w:pos="720"/></w:tabs>' +
        "</w:pPr></w:style>" +
        '<w:style w:type="paragraph" w:styleId="Body"><w:pPr>' +
        '<w:jc w:val="center"/><w:tabs>' +
        '<w:tab w:val="clear" w:pos="1080"/>' +
        '<w:tab w:val="center" w:pos="1800"/>' +
        "</w:tabs></w:pPr></w:style>",
      { numbering: NUMBERING }
    );

    const paragraph = resolveParagraph(
      `<w:pPr><w:pStyle w:val="Body"/>${LIST_PPR}<w:jc w:val="right"/>` +
        '<w:tabs><w:tab w:val="end" w:pos="2160"/></w:tabs></w:pPr>',
      formatting,
      { tableStyleId: "Grid", conditions: [] }
    );

    expect(paragraph.format).toEqual({
      spaceBeforePt: 1,
      spaceAfterPt: 2,
      align: "right",
      numbering: { numId: 4, ilvl: 2 },
      tabStops: [
        { positionPt: 90, align: "center" },
        { positionPt: 108, align: "end" },
      ],
    });
    expect(paragraph.tabStops).toEqual([
      { positionPt: 90, align: "center", layer: "style" },
      { positionPt: 108, align: "end", layer: "direct" },
    ]);
  });

  it("a paragraph pointing at no style wears the default paragraph style", () => {
    const formatting = context(
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
        '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Quote">' +
        '<w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
    );

    expect(resolveParagraph(null, formatting)).toEqual({
      format: { align: "center" },
      styleRun: { bold: true },
      inheritedRun: { bold: true },
      tabStops: [],
    });
    // Naming another style takes the default one's place rather than adding to it
    expect(
      resolveParagraph('<w:pPr><w:pStyle w:val="Quote"/></w:pPr>', formatting)
    ).toEqual({
      format: { indentStartPt: 36 },
      styleRun: null,
      inheritedRun: {},
      tabStops: [],
    });
  });

  it("lets numId zero clear numbering inherited from a style", () => {
    const formatting = context(
      '<w:style w:type="paragraph" w:styleId="Listed"><w:pPr>' +
        `${LIST_PPR}</w:pPr></w:style>`,
      { numbering: NUMBERING }
    );

    expect(
      resolveParagraph(
        '<w:pPr><w:pStyle w:val="Listed"/><w:numPr>' +
          '<w:numId w:val="0"/></w:numPr></w:pPr>',
        formatting
      ).format
    ).toBeNull();
  });

  it("layers Transitional and Strict leading indents in the same slot", () => {
    expect(
      resolveParagraph('<w:pPr><w:ind w:left="720"/></w:pPr>', {
        ...NO_FORMATTING,
        paragraphDefaults: { indentStartPt: 18 },
      }).format
    ).toEqual({ indentStartPt: 36 });
  });

  it("resolves the same paragraph to the same values wherever it stands", () => {
    const formatting = context(
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
        '<w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:b/></w:rPr>' +
        "</w:style>"
    );
    const pPr = '<w:pPr><w:jc w:val="right"/></w:pPr>';

    expect(resolveParagraph(pPr, formatting)).toEqual(
      resolveParagraph(pPr, formatting)
    );
    expect(resolveParagraph(pPr, formatting)).not.toBe(
      resolveParagraph(pPr, formatting)
    );
  });
});

describe("the run hierarchy", () => {
  const formatting = context(
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
      '<w:sz w:val="20"/><w:color w:val="333333"/>' +
      "</w:rPr></w:rPrDefault></w:docDefaults>" +
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
      '<w:rPr><w:b/><w:color w:val="2E74B5"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Strong">' +
      '<w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr></w:style>'
  );
  const paragraph: ResolvedParagraph = resolveParagraph(null, formatting);

  it("lays the paragraph style under the run and keeps docDefaults out of the mark", () => {
    expect(paragraph.styleRun).toEqual({ bold: true, color: "#2E74B5" });
    expect(resolveRun(null, paragraph, formatting)).toEqual({
      bold: true,
      color: "#2E74B5",
    });
    expect(
      resolveRun('<w:rPr><w:sz w:val="24"/></w:rPr>', paragraph, formatting)
    ).toEqual({ bold: true, color: "#2E74B5", fontSizePt: 12 });
  });

  it("resolveRun lays w:rStyle values between the paragraph style and the rPr", () => {
    expect(
      resolveRun(
        '<w:rPr><w:rStyle w:val="Strong"/></w:rPr>',
        paragraph,
        formatting
      )
    ).toEqual({ bold: true, italic: true, color: "#FF0000" });
    expect(
      resolveRun(
        '<w:rPr><w:rStyle w:val="Strong"/><w:color w:val="00FF00"/></w:rPr>',
        paragraph,
        formatting
      )
    ).toEqual({ bold: true, italic: true, color: "#00FF00" });
  });

  it("inheritedRun includes docDefaults rPr and the character style", () => {
    expect(paragraph.inheritedRun).toEqual({
      fontSizePt: 10,
      bold: true,
      color: "#2E74B5",
    });
    expect(
      inheritedRunFormat(
        '<w:rPr><w:rStyle w:val="Strong"/><w:color w:val="00FF00"/></w:rPr>',
        paragraph,
        formatting
      )
    ).toEqual({ fontSizePt: 10, bold: true, italic: true, color: "#FF0000" });
  });

  it("a run under no style and with no formatting of its own resolves to nothing", () => {
    const bare = resolveParagraph(null, NO_FORMATTING);
    expect(resolveRun(null, bare, NO_FORMATTING)).toBeNull();
    expect(inheritedRunFormat(null, bare, NO_FORMATTING)).toEqual({});
  });
});

describe("tab stops", () => {
  it("adds the implicit stop created by a hanging indent", () => {
    expect(
      resolveParagraph(
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>',
        NO_FORMATTING
      )
    ).toEqual({
      format: {
        indentStartPt: 36,
        textIndentPt: -18,
        tabStops: [{ positionPt: 36, align: "start" }],
      },
      styleRun: null,
      inheritedRun: {},
      tabStops: [{ positionPt: 36, align: "start", layer: "implicit" }],
    });
    expect(
      resolveParagraph(
        '<w:pPr><w:bidi/><w:ind w:start="720" w:hanging="360"/></w:pPr>',
        NO_FORMATTING
      ).format
    ).toEqual({
      direction: "rtl",
      indentStartPt: 36,
      textIndentPt: -18,
      tabStops: [{ positionPt: 36, align: "start" }],
    });
  });

  it("a stop written at the hanging indent's position takes the implicit one's place", () => {
    expect(
      resolveParagraph(
        '<w:pPr><w:ind w:left="720" w:hanging="360"/>' +
          '<w:tabs><w:tab w:val="center" w:pos="720"/></w:tabs></w:pPr>',
        NO_FORMATTING
      ).tabStops
    ).toEqual([{ positionPt: 36, align: "center", layer: "direct" }]);
  });

  it("creates the implicit stop from the list level's hanging indent where the paragraph writes none", () => {
    const numbering: Numbering = {
      lists: new Map([
        [
          1,
          {
            levels: new Map([
              [
                0,
                {
                  format: "decimal",
                  text: "%1.",
                  start: 1,
                  indent: {
                    startTwips: 1440,
                    endTwips: null,
                    hangingTwips: 360,
                    firstLineTwips: null,
                  },
                },
              ],
            ]),
          },
        ],
      ]),
    };
    const formatting = { ...NO_FORMATTING, numbering };
    const listed = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';

    expect(
      resolveParagraph(`<w:pPr>${listed}</w:pPr>`, formatting).tabStops
    ).toEqual([{ positionPt: 72, align: "start", layer: "implicit" }]);
    // The paragraph's own indent is the one drawn, so the stop follows it and not the level
    expect(
      resolveParagraph(
        `<w:pPr>${listed}<w:ind w:left="720" w:hanging="360"/></w:pPr>`,
        formatting
      ).tabStops
    ).toEqual([{ positionPt: 36, align: "start", layer: "implicit" }]);
    expect(
      resolveParagraph(
        `<w:pPr>${listed}<w:ind w:left="720"/></w:pPr>`,
        formatting
      ).tabStops
    ).toEqual([{ positionPt: 36, align: "start", layer: "implicit" }]);
    expect(
      resolveParagraph(
        `<w:pPr>${listed}<w:ind w:left="720" w:firstLine="360"/></w:pPr>`,
        formatting
      ).tabStops
    ).toEqual([]);
  });

  it("noTabHangInd switches the implicit stop off", () => {
    const paragraph = resolveParagraph(
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>',
      { ...NO_FORMATTING, compat: { noTabHangInd: true } }
    );
    expect(paragraph.tabStops).toEqual([]);
    expect(paragraph.format).toEqual({ indentStartPt: 36, textIndentPt: -18 });
  });

  it("each effective stop knows the layer that laid it down", () => {
    const formatting = context(
      '<w:style w:type="paragraph" w:styleId="Body"><w:pPr>' +
        '<w:tabs><w:tab w:val="center" w:pos="1800"/></w:tabs>' +
        "</w:pPr></w:style>",
      {
        numbering: NUMBERING,
        paragraphDefaults: {
          tabStops: [
            { positionPt: 36, align: "start" },
            { positionPt: 72, align: "start" },
          ],
        },
      }
    );
    const paragraph = resolveParagraph(
      `<w:pPr><w:pStyle w:val="Body"/>${LIST_PPR}` +
        '<w:ind w:left="2880" w:hanging="360"/>' +
        '<w:tabs><w:tab w:val="end" w:pos="2160"/></w:tabs></w:pPr>',
      formatting
    );

    expect(paragraph.tabStops).toEqual([
      { positionPt: 54, align: "num", layer: "numbering" },
      { positionPt: 72, align: "start", layer: "style" },
      { positionPt: 90, align: "center", layer: "style" },
      { positionPt: 108, align: "end", layer: "direct" },
      { positionPt: 144, align: "start", layer: "implicit" },
    ]);
  });

  it("the attrs carry the merged stops without their layer", () => {
    const paragraph = resolveParagraph(
      '<w:pPr><w:ind w:left="720" w:hanging="360"/>' +
        '<w:tabs><w:tab w:val="end" w:pos="2160" w:leader="dot"/></w:tabs>' +
        "</w:pPr>",
      NO_FORMATTING
    );

    expect(paragraph.format?.tabStops).toStrictEqual([
      { positionPt: 36, align: "start" },
      { positionPt: 108, align: "end", leader: "dot" },
    ]);
    for (const stop of paragraph.format?.tabStops ?? []) {
      expect(stop).not.toHaveProperty("layer");
    }
  });
});

/** Every paragraph's attrs against what the resolver answers for its own pPr. The count says the walk saw some */
function expectResolvedAttrs(
  doc: PMNode,
  formatting: FormattingContext
): number {
  let paragraphs = 0;
  doc.descendants((node) => {
    if (node.type !== docxSchema.nodes.paragraph) return true;
    paragraphs += 1;
    const pPr: unknown = node.attrs.pPr;
    const attrs = paragraphAttrsFor(
      typeof pPr === "string" ? pPr : null,
      formatting
    );
    expect(node.attrs.format).toEqual(attrs.format);
    expect(node.attrs.styleRun).toEqual(attrs.styleRun);
    return false;
  });
  return paragraphs;
}

describe("the values an opened document carries", () => {
  it.each(fixtureNames)(
    "every paragraph of %s carries what the resolver answers for its pPr",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      expect(expectResolvedAttrs(doc, session.formatting)).toBeGreaterThan(0);
    }
  );

  it("the same pPr resolves to the same attrs whichever caller asks", () => {
    const { doc, session } = importDocx(
      makeStyledNumberedDocx(
        '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>',
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
          '<w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr>' +
          "<w:rPr><w:b/></w:rPr></w:style>" +
          '<w:style w:type="paragraph" w:styleId="Heading1">' +
          '<w:name w:val="heading 1"/><w:pPr><w:jc w:val="center"/></w:pPr>' +
          "<w:rPr><w:i/></w:rPr></w:style>"
      )
    );
    const formatting = session.formatting;
    const opened = editorStateForSession({ doc, session });
    expect(expectResolvedAttrs(opened.doc, formatting)).toBe(1);

    // The style writer
    const styled = runCommand(select(opened, 1), setParagraphStyle("Heading1"));
    expect(expectResolvedAttrs(styled.doc, formatting)).toBe(1);

    // The plugin, over the paragraph Enter opened
    const split = runCommand(select(styled, 5), docxKeymap.Enter);
    expect(expectResolvedAttrs(split.doc, formatting)).toBe(2);

    // The paragraph writer every other paragraph edit goes through
    const listed = runCommand(select(split, 1), toggleNumberedList);
    expect(expectResolvedAttrs(listed.doc, formatting)).toBe(2);
    expect(toParagraphFormat(listed.doc.child(0).attrs.format)).toEqual({
      align: "center",
      numbering: { numId: 2, ilvl: 0 },
    });
  });
});
