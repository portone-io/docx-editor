// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx } from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toRunFormat } from "../../model/format";
import { createEditorState } from "../createEditor";
import { documentDefaults } from "../documentStyles";
import {
  activeFontFamily,
  activeFontSize,
  activeTextBackground,
  activeTextColor,
  documentFontNames,
  isBoldActive,
  isItalicActive,
  isUnderlineActive,
  setFontFamily,
  setFontSize,
  setTextBackground,
  setTextColor,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from "./formattingCommands";

function paragraph(rPr: string, text: string): string {
  return `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function opened(
  body: string,
  rPrDefault?: string
): { state: EditorState; session: SessionStore } {
  const { doc, session } = importDocx(makeDocx(body, rPrDefault));
  return {
    // For the toolbar to know the effective size, the document defaults have to be in the editing state
    state: createEditorState(doc, {
      styles: session.styles,
      defaults: session.defaults,
    }),
    session,
  };
}

function caret(state: EditorState, at: number): EditorState {
  return select(state, at, at);
}

/** The first paragraph that has text, and the range of that paragraph's first text piece */
function _firstTextRange(doc: PMNode): {
  index: number;
  from: number;
  to: number;
} {
  let found: { index: number; from: number; to: number } | null = null;
  doc.forEach((block, blockOffset, index) => {
    if (found !== null || block.type.name !== "paragraph") return;
    block.forEach((child, childOffset) => {
      if (found !== null || !child.isText) return;
      const from = blockOffset + 1 + childOffset;
      found = { index, from: from + 1, to: from + child.nodeSize };
    });
  });
  if (found === null) throw new Error("no paragraph with text");
  return found;
}

describe("toggle direction", () => {
  const mixed =
    "<w:p>" +
    paragraph("<w:rPr><w:b/></w:rPr>", "ab") +
    paragraph("", "cd") +
    "</w:p>";

  it("turns the whole selection on when only part of it is on, leaving the already on piece untouched", () => {
    const { state, session } = opened(mixed);
    const all = select(state, 1, 5);
    expect(isBoldActive(all)).toBe(false);

    const bolded = runCommand(all, toggleBold);
    expect(isBoldActive(bolded)).toBe(true);

    const documentXml = documentXmlOf(bolded.doc, session);
    // The rPr of the run that was already bold stays byte for byte as in the original
    expect(documentXml).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">ab</w:t></w:r>'
    );
    expect(documentXml).toContain(
      "<w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:t xml:space="preserve">cd</w:t></w:r>'
    );
  });

  it("turns the whole selection off when all of it is on", () => {
    const { state, session } = opened(mixed);
    const bolded = runCommand(select(state, 1, 5), toggleBold);
    const plain = runCommand(bolded, toggleBold);

    expect(isBoldActive(plain)).toBe(false);
    const documentXml = documentXmlOf(plain.doc, session);
    expect(documentXml).not.toContain("<w:b/>");
    expect(documentXml).not.toContain("<w:rPr>");
    expect(documentXml).toContain('<w:t xml:space="preserve">abcd</w:t>');
  });

  it("italic, underline and strikethrough follow the same convention", () => {
    const { state, session } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const range = select(state, 1, 3);
    const italic = runCommand(range, toggleItalic);
    expect(isItalicActive(italic)).toBe(true);

    const underlined = runCommand(italic, toggleUnderline);
    expect(isUnderlineActive(underlined)).toBe(true);

    const struck = runCommand(underlined, toggleStrike);
    expect(documentXmlOf(struck.doc, session)).toContain(
      '<w:rPr><w:i/><w:iCs/><w:strike/><w:u w:val="single"/></w:rPr>'
    );

    const plain = runCommand(
      runCommand(runCommand(struck, toggleItalic), toggleUnderline),
      toggleStrike
    );
    expect(documentXmlOf(plain.doc, session)).not.toContain("<w:rPr>");
  });

  it("splits only the selected piece off into its own run", () => {
    const { state, session } = opened(
      "<w:p>" + paragraph("", "abc") + "</w:p>"
    );
    const bolded = runCommand(select(state, 2, 3), toggleBold);

    expect(documentXmlOf(bolded.doc, session)).toContain(
      '<w:r><w:t xml:space="preserve">a</w:t></w:r>' +
        "<w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:t xml:space="preserve">b</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">c</w:t></w:r>'
    );
  });
});

describe("applying formatting at the caret", () => {
  it("the formatting attaches to the text typed next", () => {
    const { state, session } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const bolded = runCommand(caret(state, 2), toggleBold);

    expect(bolded.storedMarks).not.toBeNull();
    expect(isBoldActive(bolded)).toBe(true);

    const typed = bolded.apply(bolded.tr.insertText("bold", 2));
    const documentXml = documentXmlOf(typed.doc, session);
    expect(documentXml).toContain(
      '<w:r><w:t xml:space="preserve">a</w:t></w:r>' +
        "<w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:t xml:space="preserve">bold</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">b</w:t></w:r>'
    );
  });

  it("inherits the formatting at the caret as is and changes only the one item", () => {
    const rPr = '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>';
    const { state, session } = opened(
      "<w:p>" + paragraph(rPr, "ab") + "</w:p>"
    );
    const colored = runCommand(caret(state, 2), setTextColor("#FF0000"));
    const typed = colored.apply(colored.tr.insertText("red", 2));

    expect(documentXmlOf(typed.doc, session)).toContain(
      '<w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr>' +
        '<w:t xml:space="preserve">red</w:t>'
    );
  });

  it("formatting can be armed ahead of time even in an empty paragraph", () => {
    const { state, session } = opened("<w:p/>");
    const bolded = runCommand(caret(state, 1), toggleBold);
    const typed = bolded.apply(bolded.tr.insertText("new text", 1));

    expect(documentXmlOf(typed.doc, session)).toContain(
      "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:t xml:space="preserve">new text</w:t></w:r></w:p>'
    );
  });
});

describe("size, color and background color", () => {
  it("size goes out as a pair of half point values", () => {
    const { state, session } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const sized = runCommand(select(state, 1, 3), setFontSize(14));

    expect(activeFontSize(sized)).toEqual({ kind: "size", pt: 14 });
    expect(documentXmlOf(sized.doc, session)).toContain(
      '<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    );
  });

  it("withdraws the setting when given null", () => {
    const rPr = '<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>';
    const { state, session } = opened(
      "<w:p>" + paragraph(rPr, "ab") + "</w:p>"
    );
    const cleared = runCommand(select(state, 1, 3), setFontSize(null));

    // Where the setting was withdrawn, the size actually rendered is reported. With no document default that is 10pt
    expect(activeFontSize(cleared)).toEqual({ kind: "default", pt: 10 });
    expect(documentXmlOf(cleared.doc, session)).not.toContain("<w:sz ");
  });

  it("shows the document default size where nothing is specified", () => {
    const { state } = opened(
      "<w:p>" + paragraph("", "ab") + "</w:p>",
      '<w:rPr><w:sz w:val="24"/></w:rPr>'
    );

    expect(activeFontSize(select(state, 1, 3))).toEqual({
      kind: "default",
      pt: 12,
    });
  });

  it("applies and withdraws the background color and the text color", () => {
    const { state, session } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const marked = runCommand(
      select(state, 1, 3),
      setTextBackground("#FFFF00")
    );
    expect(activeTextBackground(marked)).toBe("#FFFF00");

    const colored = runCommand(marked, setTextColor("#2E74B5"));
    expect(activeTextColor(colored)).toBe("#2E74B5");
    expect(documentXmlOf(colored.doc, session)).toContain(
      '<w:rPr><w:color w:val="2E74B5"/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:rPr>'
    );

    const plain = runCommand(
      runCommand(colored, setTextBackground(null)),
      setTextColor(null)
    );
    expect(documentXmlOf(plain.doc, session)).not.toContain("<w:rPr>");
  });

  it("painting a color over an old highlight document turns the highlight into shading", () => {
    const { state, session } = opened(
      "<w:p>" +
        paragraph('<w:rPr><w:highlight w:val="yellow"/></w:rPr>', "ab") +
        "</w:p>"
    );
    // A highlight written as a name is also reported as a color so the palette can compare against it
    expect(activeTextBackground(select(state, 1, 3))).toBe("#ffff00");

    const repainted = runCommand(
      select(state, 1, 3),
      setTextBackground("#FF0000")
    );
    const xml = documentXmlOf(repainted.doc, session);
    expect(xml).not.toContain("<w:highlight");
    expect(xml).toContain(
      '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>'
    );
    expect(activeTextBackground(repainted)).toBe("#FF0000");
  });

  it("withdrawing an old highlight with None makes the highlight disappear", () => {
    const { state, session } = opened(
      "<w:p>" +
        paragraph('<w:rPr><w:highlight w:val="cyan"/></w:rPr>', "ab") +
        "</w:p>"
    );
    const plain = runCommand(select(state, 1, 3), setTextBackground(null));

    expect(documentXmlOf(plain.doc, session)).not.toContain("<w:highlight");
    expect(activeTextBackground(plain)).toBeNull();
  });

  it("a selection mixing several values has no active value", () => {
    const body =
      "<w:p>" +
      paragraph('<w:rPr><w:sz w:val="20"/></w:rPr>', "ab") +
      paragraph('<w:rPr><w:sz w:val="24"/></w:rPr>', "cd") +
      "</w:p>";
    const { state } = opened(body);
    // A mixed selection and "nothing specified" are different states. Only the mixed one has no number to show
    expect(activeFontSize(select(state, 1, 5))).toEqual({ kind: "mixed" });
    expect(activeFontSize(select(state, 1, 3))).toEqual({
      kind: "size",
      pt: 10,
    });
  });

  it("answers that it has nothing to do for a value that cannot be written to the document", () => {
    const { state } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const range = select(state, 1, 3);
    expect(setFontSize(0.3)(range)).toBe(false);
    expect(setTextColor("red")(range)).toBe(false);
    // Formatting already in that state also has nothing to do
    expect(setTextColor(null)(range)).toBe(false);
  });
});

describe("font", () => {
  /** The shape in which the fixtures write fonts. The same name appears in all four slots */
  const rFonts = (name: string) =>
    `<w:rPr><w:rFonts w:ascii="${name}" w:cs="${name}" ` +
    `w:eastAsia="${name}" w:hAnsi="${name}"/></w:rPr>`;

  it("changes the font of the selected text", () => {
    const { state, session } = opened(
      "<w:p>" + paragraph(rFonts("Malgun Gothic"), "ab") + "</w:p>"
    );
    const changed = runCommand(select(state, 1, 3), setFontFamily("Batang"));

    expect(activeFontFamily(changed)).toEqual({ kind: "font", name: "Batang" });
    expect(documentXmlOf(changed.doc, session)).toContain(
      '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" ' +
        'w:eastAsia="Batang" w:cs="Batang"/>'
    );
  });

  it("withdraws the setting when given null and shows the font that actually gets rendered", () => {
    const { state, session } = opened(
      "<w:p>" + paragraph(rFonts("Batang"), "ab") + "</w:p>",
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
    );
    const cleared = runCommand(select(state, 1, 3), setFontFamily(null));

    expect(activeFontFamily(cleared)).toEqual({
      kind: "default",
      name: "Arial",
    });
    expect(documentXmlOf(cleared.doc, session)).not.toContain("<w:rFonts");
  });

  it("reports the fallback font name when the document did not write down a default font either", () => {
    const { state } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    expect(activeFontFamily(select(state, 1, 3))).toEqual({
      kind: "default",
      name: "Arial",
    });
  });

  it("a selection mixing fonts has no active value", () => {
    const body =
      "<w:p>" +
      paragraph(rFonts("Batang"), "ab") +
      paragraph(rFonts("Arial"), "cd") +
      "</w:p>";
    const { state } = opened(body);

    expect(activeFontFamily(select(state, 1, 5))).toEqual({ kind: "mixed" });
    expect(activeFontFamily(select(state, 1, 3))).toEqual({
      kind: "font",
      name: "Batang",
    });
  });

  it("collects the fonts the document uses", () => {
    const body =
      "<w:p>" +
      paragraph(rFonts("Malgun Gothic"), "ab") +
      paragraph(rFonts("Batang"), "cd") +
      paragraph(rFonts("Malgun Gothic"), "ef") +
      "</w:p>";
    const { state } = opened(
      body,
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
    );

    // The document default font is included too, and the same name appears only once
    expect(documentFontNames(state.doc, documentDefaults(state))).toEqual([
      "Arial",
      "Batang",
      "Malgun Gothic",
    ]);
  });

  it("answers that it has nothing to do for a name that cannot be written to the document", () => {
    const { state } = opened("<w:p>" + paragraph("", "ab") + "</w:p>");
    const range = select(state, 1, 3);
    expect(setFontFamily("")(range)).toBe(false);
    expect(setFontFamily('Ari"al')(range)).toBe(false);
    // Withdrawing where there is already no setting also has nothing to do
    expect(setFontFamily(null)(range)).toBe(false);
  });
});

describe("a selection with things other than text mixed in", () => {
  const body =
    "<w:p>" +
    '<w:bookmarkStart w:id="0" w:name="start"/>' +
    '<w:r><w:t xml:space="preserve">ab</w:t><w:br/><w:tab/>' +
    '<w:t xml:space="preserve">cd</w:t></w:r>' +
    '<w:bookmarkEnd w:id="0"/>' +
    "</w:p>";

  it("formats text and tabs while passing over preserved elements and line breaks", () => {
    const { state, session } = opened(body);
    const whole = select(state, 1, state.doc.child(0).nodeSize - 1);
    const bolded = runCommand(whole, toggleBold);

    const documentXml = documentXmlOf(bolded.doc, session);
    // The bookmarks stay in place byte for byte as in the original
    expect(documentXml).toContain('<w:bookmarkStart w:id="0" w:name="start"/>');
    expect(documentXml).toContain('<w:bookmarkEnd w:id="0"/>');
    expect(documentXml).toContain(
      "<w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:t xml:space="preserve">ab</w:t></w:r>' +
        "<w:r><w:br/></w:r>" +
        "<w:r><w:rPr><w:b/><w:bCs/></w:rPr>" +
        '<w:tab/><w:t xml:space="preserve">cd</w:t></w:r>'
    );
  });

  it("writes a selected tab's font size into its OOXML run", () => {
    const { state, session } = opened("<w:p><w:r><w:tab/></w:r></w:p>");
    const sized = runCommand(select(state, 1, 2), setFontSize(36));

    expect(documentXmlOf(sized.doc, session)).toContain(
      '<w:r><w:rPr><w:sz w:val="72"/><w:szCs w:val="72"/></w:rPr><w:tab/></w:r>'
    );
  });
});

describe("table cell selection", () => {
  const cell = (text: string) =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
  const body =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cell("Left")}${cell("Right")}</w:tr>` +
    "</w:tbl>";

  it("changes the text of all the selected cells together", () => {
    const { state, session } = opened(body);
    const cells: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type.name === "tableCell") cells.push(pos);
      return true;
    });

    const selected = state.apply(
      state.tr.setSelection(CellSelection.create(state.doc, cells[0], cells[1]))
    );
    const bolded = runCommand(selected, toggleBold);

    const documentXml = documentXmlOf(bolded.doc, session);
    expect(documentXml.match(/<w:b\/>/g)).toHaveLength(2);
    expect(isBoldActive(bolded)).toBe(true);
  });
});

describe("round trip", () => {
  it("an edited paragraph only gains and loses formatting items", () => {
    const rPr =
      '<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>' +
      '<w:szCs w:val="20"/></w:rPr>';
    const { state, session } = opened(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        paragraph(rPr, "ab") +
        "</w:p>"
    );
    const bolded = runCommand(select(state, 1, 3), toggleBold);

    expect(documentXmlOf(bolded.doc, session)).toContain(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:b/><w:bCs/>' +
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>' +
        '<w:t xml:space="preserve">ab</w:t></w:r></w:p>'
    );
  });

  it("reopening a document exported with formatting applied reads that formatting back unchanged", () => {
    const { state, session } = opened(
      "<w:p>" + paragraph("", "abc") + "</w:p>"
    );
    const edited = runCommand(
      runCommand(select(state, 1, 3), toggleBold),
      setTextBackground("#00FFFF")
    );

    const reopened = importDocx(exportDocx(edited.doc, session));
    const first = reopened.doc.child(0).child(0);
    expect(first.text).toBe("ab");
    expect(toRunFormat(first.marks[0].attrs.format)).toEqual({
      bold: true,
      background: "#00FFFF",
    });
    expect(reopened.doc.child(0).child(1).marks[0].attrs.rPr).toBeNull();
  });
});

describe("formatting a selection that runs over locked text", () => {
  const LOCKED_PR =
    '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

  /** Plain text, a locked control, then plain text again */
  const WITH_LOCK =
    "<w:p>" +
    paragraph("", "a") +
    `<w:sdt>${LOCKED_PR}<w:sdtContent>${paragraph("", "bc")}` +
    "</w:sdtContent></w:sdt>" +
    paragraph("", "d") +
    "</w:p>";

  // Asking for the locked part as well would have the guard turn the whole transaction down
  it("applies to every stretch the lock leaves open", () => {
    const { state } = opened(WITH_LOCK);
    const bolded = runCommand(select(state, 1, 5), toggleBold);

    expect(isBoldActive(select(bolded, 1, 2))).toBe(true);
    expect(isBoldActive(select(bolded, 2, 4))).toBe(false);
    expect(isBoldActive(select(bolded, 4, 5))).toBe(true);
  });

  it("has nothing to do where the selection holds locked text alone", () => {
    const { state } = opened(WITH_LOCK);
    expect(toggleBold(select(state, 2, 4))).toBe(false);
  });

  // Counting the locked stretch would leave it reading off for text that is already on
  it("reads as on once every stretch it can reach is on", () => {
    const { state } = opened(WITH_LOCK);
    const bolded = runCommand(select(state, 1, 5), toggleBold);

    expect(isBoldActive(select(bolded, 1, 5))).toBe(true);
  });

  it("takes the format off again on a second press", () => {
    const { state } = opened(WITH_LOCK);
    const bolded = runCommand(select(state, 1, 5), toggleBold);
    const cleared = runCommand(select(bolded, 1, 5), toggleBold);

    expect(isBoldActive(select(cleared, 1, 2))).toBe(false);
    expect(isBoldActive(select(cleared, 4, 5))).toBe(false);
  });
});
