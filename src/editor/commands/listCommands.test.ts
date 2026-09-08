// @vitest-environment jsdom

import { undo } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledNumberedDocx,
} from "../../__testing__/docx";
import { posOfText, runCommand, select } from "../../__testing__/editing";
import { NO_FORMATTING } from "../../docx/formatting";
import { importDocx } from "../../docx/importDocx";
import { toParagraphFormat } from "../../model/format";
import { newListsOf } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import {
  EMPTY_NUMBERING,
  type Numbering,
  parseNumbering,
} from "../../numbering/parseNumbering";
import { createEditorState, editorStateForSession } from "../createEditor";
import { type EditorDocument, NO_DOCUMENT } from "../editorDocument";
import { docxKeymap } from "../plugins/keymap";
import {
  documentNumbering,
  paragraphMarkers,
} from "../plugins/numberingDecorations";
import {
  activeListKind,
  decreaseListLevel,
  increaseListLevel,
  removeFromList,
  toggleBulletList,
  toggleNumberedList,
} from "./listCommands";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** When `left` is null, that level specifies no indentation */
function level(ilvl: number, left: number | null): string {
  const ind =
    left === null
      ? ""
      : `<w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>`;
  return (
    `<w:lvl w:ilvl="${ilvl}"><w:numFmt w:val="decimal"/>` +
    `<w:lvlText w:val="%${ilvl + 1}."/>${ind}</w:lvl>`
  );
}

/** A three-level list definition. The third level specifies no indentation */
function threeLevelNumbering(): Numbering {
  return parseNumbering(
    `<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0">` +
      level(0, 720) +
      level(1, 1440) +
      level(2, null) +
      "</w:abstractNum>" +
      '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>' +
      "</w:numbering>"
  );
}

function paragraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function listPPr(numId: number, ilvl: number, ind: string): string {
  return (
    `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/>` +
    `<w:numId w:val="${numId}"/></w:numPr>${ind}</w:pPr>`
  );
}

/** A state over list definitions the test writes, which the body points into */
function openState(body: string, numbering = EMPTY_NUMBERING): EditorState {
  const { doc } = importDocx(makeDocx(body));
  const withLists: EditorDocument = {
    ...NO_DOCUMENT,
    formatting: { ...NO_FORMATTING, numbering },
  };
  return createEditorState(doc, { document: withLists });
}

function paragraphAt(doc: PMNode, index: number): PMNode {
  return doc.child(index);
}

function listRef(node: PMNode) {
  return toParagraphFormat(node.attrs.format)?.numbering ?? null;
}

function pPrOf(node: PMNode): string {
  const pPr: unknown = node.attrs.pPr;
  return typeof pPr === "string" ? pPr : "";
}

/** The numbers drawn on screen, in document order */
function markerTexts(doc: PMNode, numbering: Numbering): string[] {
  return paragraphMarkers(doc, numbering).map((marker) => marker.text);
}

describe("moving a list level", () => {
  const body = paragraph(
    "Entry",
    listPPr(3, 0, '<w:ind w:left="720" w:hanging="360"/>')
  );

  function firstItem(state: EditorState): EditorState {
    return select(state, posOfText(state.doc, "Entry"));
  }

  it("Tab takes the level down one step and applies that level's indentation", () => {
    const state = firstItem(openState(body, threeLevelNumbering()));
    const next = runCommand(state, increaseListLevel);
    const item = paragraphAt(next.doc, 0);

    expect(listRef(item)).toEqual({ numId: 3, ilvl: 1 });
    expect(pPrOf(item)).toContain('<w:ind w:left="1440" w:hanging="360"/>');
    expect(toParagraphFormat(item.attrs.format)?.indentStartPt).toBe(72);
  });

  it("Shift+Tab takes the level up one step", () => {
    const deep = paragraph("Entry", listPPr(3, 2, '<w:ind w:left="2160"/>'));
    const state = firstItem(openState(deep, threeLevelNumbering()));
    const item = paragraphAt(runCommand(state, decreaseListLevel).doc, 0);

    expect(listRef(item)).toEqual({ numId: 3, ilvl: 1 });
    expect(pPrOf(item)).toContain('<w:ind w:left="1440" w:hanging="360"/>');
  });

  it("does not go any further up from the first level", () => {
    const state = firstItem(openState(body, threeLevelNumbering()));
    expect(decreaseListLevel(state)).toBe(false);
  });

  it("does not go any further down from the last level", () => {
    const deepest = paragraph("Entry", listPPr(3, 8, ""));
    const state = firstItem(openState(deepest, threeLevelNumbering()));
    expect(increaseListLevel(state)).toBe(false);
  });

  it("leaves the paragraph's value alone when moving to a level whose definition has no indentation", () => {
    const second = paragraph(
      "Entry",
      listPPr(3, 1, '<w:ind w:left="1440" w:hanging="360"/>')
    );
    const state = firstItem(openState(second, threeLevelNumbering()));
    const item = paragraphAt(runCommand(state, increaseListLevel).doc, 0);

    expect(listRef(item)).toEqual({ numId: 3, ilvl: 2 });
    expect(pPrOf(item)).toContain('<w:ind w:left="1440" w:hanging="360"/>');
  });

  it("a paragraph that records no indentation is moved by the level alone", () => {
    const numbering = threeLevelNumbering();
    const state = openState(paragraph("Item", listPPr(3, 0, "")), numbering);
    const moved = runCommand(
      select(state, posOfText(state.doc, "Item")),
      increaseListLevel
    );
    const item = paragraphAt(moved.doc, 0);

    expect(listRef(item)).toEqual({ numId: 3, ilvl: 1 });
    expect(pPrOf(item)).not.toContain("<w:ind");
    // The second level of the definition indents by 1440 twips, drawn by the marker decoration
    expect(paragraphMarkers(moved.doc, numbering)[0].indentStartPt).toBe(72);
  });

  it("updates a Strict leading indent when the list level changes", () => {
    const numbering = threeLevelNumbering();
    const body = paragraph(
      "Item",
      listPPr(3, 0, '<w:ind w:start="720" w:hanging="360"/>')
    );
    const state = openState(body, numbering);
    const moved = runCommand(
      select(state, posOfText(state.doc, "Item")),
      increaseListLevel
    );
    const item = paragraphAt(moved.doc, 0);

    expect(listRef(item)).toEqual({ numId: 3, ilvl: 1 });
    expect(toParagraphFormat(item.attrs.format)?.indentStartPt).toBe(72);
    expect(pPrOf(item)).toContain('<w:ind w:left="1440" w:hanging="360"/>');
  });

  it("does nothing in a paragraph that is not a list", () => {
    const state = firstItem(openState(paragraph("Entry")));
    expect(increaseListLevel(state)).toBe(false);
    expect(decreaseListLevel(state)).toBe(false);
  });

  it("all the selected paragraphs move together", () => {
    const two = body + paragraph("Next", listPPr(3, 0, ""));
    const state = openState(two, threeLevelNumbering());
    const both = select(
      state,
      posOfText(state.doc, "Entry"),
      posOfText(state.doc, "Next")
    );
    const moved = runCommand(both, increaseListLevel);

    expect(listRef(paragraphAt(moved.doc, 0))?.ilvl).toBe(1);
    expect(listRef(paragraphAt(moved.doc, 1))?.ilvl).toBe(1);
  });
});

describe("starting a new list", () => {
  const body = paragraph("First") + paragraph("Second");

  function bothParagraphs(state: EditorState): EditorState {
    return select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
  }

  it("paragraphs selected in one go use the same numbering id", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const listed = runCommand(state, toggleNumberedList);
    const first = listRef(paragraphAt(listed.doc, 0));
    const second = listRef(paragraphAt(listed.doc, 1));

    expect(first).toEqual(second);
    expect(first?.ilvl).toBe(0);
  });

  it("picks a new numbering id, avoiding the ones already in use", () => {
    const used = paragraph("Already", listPPr(4, 0, "")) + body;
    const state = openState(used, threeLevelNumbering());
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    const numId = listRef(
      paragraphAt(runCommand(both, toggleNumberedList).doc, 1)
    )?.numId;

    // It goes past 3 from the definition and 4 used by a paragraph
    expect(numId).toBeGreaterThan(4);
  });

  it("the kind decides the definition and not the number, so both kinds take the same one", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const numbered = runCommand(state, toggleNumberedList);
    const bulleted = runCommand(state, toggleBulletList);

    // Started from the same document, each takes the first number nothing else spends
    expect(listRef(paragraphAt(numbered.doc, 0))?.numId).toBe(4);
    expect(listRef(paragraphAt(bulleted.doc, 0))?.numId).toBe(4);
    expect(newListsOf(numbered.doc.attrs.newLists).get(4)).toEqual(
      templateList("numbered")
    );
    expect(newListsOf(bulleted.doc.attrs.newLists).get(4)).toEqual(
      templateList("bullet")
    );
  });

  it("two lists started one after the other take different numbers and the same definition", () => {
    const state = openState(body, threeLevelNumbering());
    const first = runCommand(
      select(state, posOfText(state.doc, "First")),
      toggleNumberedList
    );
    const second = runCommand(
      select(first, posOfText(first.doc, "Second")),
      toggleNumberedList
    );
    const registered = newListsOf(second.doc.attrs.newLists);

    expect([...registered.keys()]).toEqual([4, 5]);
    expect(registered.get(4)).toEqual(registered.get(5));
  });

  it("records the definition the list was started with on the document node", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const listed = runCommand(state, toggleBulletList);
    const numId = listRef(paragraphAt(listed.doc, 0))?.numId;
    const registered = newListsOf(listed.doc.attrs.newLists);

    expect(numId).toBeDefined();
    expect([...registered.keys()]).toEqual([numId]);
    // The definition is the one a bullet list is started with, whatever number it was given
    expect(registered.get(numId ?? 0)).toEqual(templateList("bullet"));
  });

  it("draws the list from the definition it registered rather than from its number", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const listed = runCommand(state, toggleBulletList);

    expect(markerTexts(listed.doc, documentNumbering(listed))).toEqual([
      "●",
      "●",
    ]);
  });

  it("undo takes the definition back off with the list", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const listed = runCommand(state, toggleNumberedList);
    expect(newListsOf(listed.doc.attrs.newLists).size).toBe(1);

    const back = runCommand(listed, undo);

    expect(newListsOf(back.doc.attrs.newLists).size).toBe(0);
    expect(back.doc.eq(state.doc)).toBe(true);
  });

  it("leaving the list gives back the number and the definition it took", () => {
    const state = bothParagraphs(openState(body, threeLevelNumbering()));
    const listed = runCommand(state, toggleNumberedList);
    const plain = runCommand(bothParagraphs(listed), removeFromList);

    expect(newListsOf(plain.doc.attrs.newLists).size).toBe(0);
    const again = runCommand(bothParagraphs(plain), toggleNumberedList);
    expect(listRef(paragraphAt(again.doc, 0))).toEqual(
      listRef(paragraphAt(listed.doc, 0))
    );
  });

  it("a new list paragraph records no indentation and takes the level's on screen", () => {
    const numbering = threeLevelNumbering();
    const state = bothParagraphs(openState(body, numbering));
    const listed = runCommand(state, toggleNumberedList);

    expect(pPrOf(paragraphAt(listed.doc, 0))).not.toContain("<w:ind");
    // The started list's first level indents by 720 twips, drawn by the marker decoration
    expect(
      paragraphMarkers(listed.doc, documentNumbering(listed))[0]
    ).toMatchObject({
      indentStartPt: 36,
      textIndentPt: -18,
    });
  });
});

describe("toggling the list buttons", () => {
  /**
   * Two lists whose definitions contradict the odd/even rule of their numbering ids.
   * Id 3 (odd) is a numbered list and id 2 (even) is a bullet list.
   */
  function mismatchedNumbering(): Numbering {
    return parseNumbering(
      `<w:numbering ${W_NS}>` +
        `<w:abstractNum w:abstractNumId="0">${level(0, 720)}</w:abstractNum>` +
        '<w:abstractNum w:abstractNumId="1">' +
        '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
        '<w:lvlText w:val="●"/></w:lvl></w:abstractNum>' +
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>' +
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
        "</w:numbering>"
    );
  }

  it("starts a list of that kind when the paragraph is not a list", () => {
    const state = openState(paragraph("Entry"), threeLevelNumbering());
    const listed = runCommand(
      select(state, posOfText(state.doc, "Entry")),
      toggleBulletList
    );

    expect(activeListKind(select(listed, posOfText(listed.doc, "Entry")))).toBe(
      "bullet"
    );
  });

  it("leaves the list when it is already of that kind", () => {
    const body = paragraph("Entry", listPPr(3, 0, ""));
    const state = openState(body, threeLevelNumbering());
    const plain = runCommand(
      select(state, posOfText(state.doc, "Entry")),
      toggleNumberedList
    );

    expect(listRef(paragraphAt(plain.doc, 0))).toBeNull();
  });

  it("switches to that kind when it is of a different kind", () => {
    const body = paragraph("Entry", listPPr(2, 0, ""));
    const state = openState(body, mismatchedNumbering());
    const numbered = runCommand(
      select(state, posOfText(state.doc, "Entry")),
      toggleNumberedList
    );
    const item = paragraphAt(numbered.doc, 0);

    expect(listRef(item)?.numId).not.toBe(2);
    expect(
      activeListKind(select(numbered, posOfText(numbered.doc, "Entry")))
    ).toBe("numbered");
  });

  it("the document's definition decides the kind, and the numbering id decides it when there is no definition", () => {
    const body =
      paragraph("Numbered", listPPr(3, 0, "")) +
      paragraph("Bullet", listPPr(2, 0, ""));
    const state = openState(body, mismatchedNumbering());

    // By the odd/even rule alone these would read the other way around
    expect(
      activeListKind(select(state, posOfText(state.doc, "Numbered")))
    ).toBe("numbered");
    expect(activeListKind(select(state, posOfText(state.doc, "Bullet")))).toBe(
      "bullet"
    );

    // A number nothing defines is a list whose kind cannot be told, so neither button is pressed
    const unknown = openState(paragraph("New", listPPr(9, 0, "")));
    expect(
      activeListKind(select(unknown, posOfText(unknown.doc, "New")))
    ).toBeNull();
  });

  it("presses no button when the paragraph is not a list or the kinds are mixed", () => {
    const state = openState(paragraph("Entry"), threeLevelNumbering());
    expect(
      activeListKind(select(state, posOfText(state.doc, "Entry")))
    ).toBeNull();

    const mixed = openState(
      paragraph("Numbered", listPPr(3, 0, "")) + paragraph("Plain"),
      mismatchedNumbering()
    );
    const both = select(
      mixed,
      posOfText(mixed.doc, "Numbered"),
      posOfText(mixed.doc, "Plain")
    );
    expect(activeListKind(both)).toBeNull();
  });
});

describe("removing from a list", () => {
  it("the list marker and the hanging indent disappear", () => {
    const body = paragraph(
      "Entry",
      listPPr(3, 0, '<w:ind w:left="720" w:hanging="360"/>')
    );
    const state = openState(body, threeLevelNumbering());
    const item = paragraphAt(
      runCommand(select(state, posOfText(state.doc, "Entry")), removeFromList)
        .doc,
      0
    );

    expect(listRef(item)).toBeNull();
    expect(pPrOf(item)).toBe('<w:pPr><w:ind w:left="720"/></w:pPr>');
  });

  it("does nothing in a paragraph that is not a list", () => {
    const state = openState(paragraph("Entry"));
    expect(removeFromList(select(state, posOfText(state.doc, "Entry")))).toBe(
      false
    );
  });
});

describe("leaving a list puts the paragraph back where it was", () => {
  const TEXT = "Body text";

  /** Runs the commands one after another, each with the caret in that paragraph */
  function runAll(
    state: EditorState,
    commands: readonly Command[]
  ): EditorState {
    return commands.reduce(
      (current, command) =>
        runCommand(select(current, posOfText(current.doc, TEXT)), command),
      state
    );
  }

  it("a paragraph that had no indentation comes back with none", () => {
    const state = openState(paragraph(TEXT), threeLevelNumbering());
    const plain = runAll(state, [toggleNumberedList, toggleNumberedList]);
    const item = paragraphAt(plain.doc, 0);

    expect(listRef(item)).toBeNull();
    expect(pPrOf(item)).toBe("");
  });

  it("an indentation the document gave the paragraph comes back", () => {
    const ind = '<w:pPr><w:ind w:left="400" w:firstLine="200"/></w:pPr>';
    const state = openState(paragraph(TEXT, ind), threeLevelNumbering());
    const plain = runAll(state, [toggleNumberedList, toggleNumberedList]);

    expect(pPrOf(paragraphAt(plain.doc, 0))).toBe(ind);
  });

  it("moving the item two levels down leaves no indentation behind", () => {
    const state = openState(paragraph(TEXT), threeLevelNumbering());
    const deep = runAll(state, [
      toggleNumberedList,
      increaseListLevel,
      increaseListLevel,
    ]);
    expect(listRef(paragraphAt(deep.doc, 0))?.ilvl).toBe(2);

    const plain = runAll(deep, [toggleNumberedList]);
    expect(listRef(paragraphAt(plain.doc, 0))).toBeNull();
    expect(pPrOf(paragraphAt(plain.doc, 0))).toBe("");
  });

  it("neither the list nor an indentation is left in the exported xml", () => {
    const { doc, session } = importDocx(makeNumberedDocx(paragraph(TEXT)));
    const state = editorStateForSession({ doc, session });
    const plain = runAll(state, [
      toggleNumberedList,
      increaseListLevel,
      toggleNumberedList,
    ]);
    const xml = documentXmlOf(plain.doc, session);

    expect(xml).toContain(TEXT);
    expect(xml).not.toContain("<w:numPr>");
    expect(xml).not.toContain("<w:ind");
  });
});

describe("a paragraph pointing at a style", () => {
  /**
   * A single-page document where the `Item` style passes down center alignment, and with a
   * numbering.xml so that a list can be started in it
   */
  function openStyled(): EditorState {
    const { doc, session } = importDocx(
      makeStyledNumberedDocx(
        paragraph("Entry", '<w:pPr><w:pStyle w:val="Item"/></w:pPr>'),
        '<w:style w:styleId="Item">' +
          '<w:pPr><w:jc w:val="center"/></w:pPr></w:style>'
      )
    );
    return editorStateForSession({ doc, session });
  }

  it("the alignment the style gives stays on screen even after starting a list", () => {
    const state = openStyled();
    expect(toParagraphFormat(state.doc.child(0).attrs.format)?.align).toBe(
      "center"
    );

    const item = paragraphAt(
      runCommand(
        select(state, posOfText(state.doc, "Entry")),
        toggleNumberedList
      ).doc,
      0
    );
    expect(listRef(item)).not.toBeNull();
    expect(toParagraphFormat(item.attrs.format)?.align).toBe("center");
  });
});

describe("Tab precedence", () => {
  const cell = (text: string, pPr = "") =>
    `<w:tc>${paragraph(text, pPr)}</w:tc>`;
  const tableBody =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cell("Left", listPPr(3, 0, ""))}${cell("Right")}</w:tr>` +
    "</w:tbl>";

  it("moving between cells wins in a list paragraph inside a table", () => {
    const state = openState(tableBody, threeLevelNumbering());
    const moved = runCommand(
      select(state, posOfText(state.doc, "Left")),
      docxKeymap.Tab
    );

    expect(moved.selection.$head.parent.textContent).toBe("Right");
    // The level is unchanged
    let ilvl: number | null = null;
    moved.doc.descendants((node) => {
      if (node.textContent === "Left" && node.type.name === "paragraph") {
        ilvl = listRef(node)?.ilvl ?? null;
      }
      return true;
    });
    expect(ilvl).toBe(0);
  });

  it("moves the level in a list paragraph outside a table", () => {
    const body = paragraph("Entry", listPPr(3, 0, ""));
    const state = openState(body, threeLevelNumbering());
    const moved = runCommand(
      select(state, posOfText(state.doc, "Entry")),
      docxKeymap.Tab
    );
    expect(listRef(paragraphAt(moved.doc, 0))?.ilvl).toBe(1);
  });

  it("inserts a document tab when it is neither a list nor a table", () => {
    const state = openState(paragraph("Entry"));
    const at = select(state, posOfText(state.doc, "Entry"));
    const inserted = runCommand(at, docxKeymap.Tab);
    expect(paragraphAt(inserted.doc, 0).textContent).toBe("E\tntry");
    expect(docxKeymap["Shift-Tab"](at)).toBe(false);
  });
});

describe("Enter in an empty list item", () => {
  /** The left indentation each level of `threeLevelNumbering` lays down, in twips */
  const LEVEL_LEFT = [720, 1440];

  function levelInd(ilvl: number): string {
    return `<w:ind w:left="${LEVEL_LEFT[ilvl]}" w:hanging="360"/>`;
  }

  /** An item the way Word writes one: the level's indentation is recorded on the paragraph itself */
  function item(text: string, ilvl = 0): string {
    return paragraph(text, listPPr(3, ilvl, levelInd(ilvl)));
  }

  /** The same item with nothing typed into it yet */
  function emptyItem(ilvl = 0): string {
    return `<w:p>${listPPr(3, ilvl, levelInd(ilvl))}</w:p>`;
  }

  /** The only caret position inside the paragraph at `index`, which holds no text */
  function insideEmpty(state: EditorState, index: number): number {
    let at = 1;
    for (let before = 0; before < index; before += 1) {
      at += state.doc.child(before).nodeSize;
    }
    return at;
  }

  it("an item at the first level becomes a plain paragraph at the start of the line", () => {
    const numbering = threeLevelNumbering();
    const state = openState(item("One") + item("Two") + emptyItem(), numbering);
    const next = runCommand(
      select(state, insideEmpty(state, 2)),
      docxKeymap.Enter
    );
    const abandoned = paragraphAt(next.doc, 2);

    // No new paragraph appears; the empty one stops being an item
    expect(next.doc.childCount).toBe(3);
    expect(listRef(abandoned)).toBeNull();
    // Nothing of its list life is left: neither the marker slot nor the level's indentation
    expect(pPrOf(abandoned)).toBe("");
    expect(toParagraphFormat(abandoned.attrs.format)?.indentStartPt).toBe(
      undefined
    );
    expect(next.selection.$from.parentOffset).toBe(0);
    // The items before it are numbered just as they were
    expect(markerTexts(next.doc, numbering)).toEqual(["1.", "2."]);
  });

  it("the left indentation goes even when the document gave it, and the rest stays", () => {
    const own = `<w:p>${listPPr(3, 0, '<w:ind w:left="400" w:right="200"/>')}</w:p>`;
    const state = openState(item("One") + own, threeLevelNumbering());
    const next = runCommand(
      select(state, insideEmpty(state, 1)),
      docxKeymap.Enter
    );

    expect(pPrOf(paragraphAt(next.doc, 1))).toBe(
      '<w:pPr><w:ind w:right="200"/></w:pPr>'
    );
  });

  it("an item in the middle leaves the items after it in the list", () => {
    const numbering = threeLevelNumbering();
    const body = item("One") + item("Two") + emptyItem() + item("Four");
    const state = openState(body, numbering);
    const next = runCommand(
      select(state, insideEmpty(state, 2)),
      docxKeymap.Enter
    );

    expect(listRef(paragraphAt(next.doc, 3))).toEqual({ numId: 3, ilvl: 0 });
    expect(markerTexts(next.doc, numbering)).toEqual(["1.", "2.", "3."]);
  });

  it("a deeper item comes up one level and stays in the list", () => {
    const state = openState(item("One") + emptyItem(1), threeLevelNumbering());
    const next = runCommand(
      select(state, insideEmpty(state, 1)),
      docxKeymap.Enter
    );
    const moved = paragraphAt(next.doc, 1);

    expect(listRef(moved)).toEqual({ numId: 3, ilvl: 0 });
    expect(pPrOf(moved)).toContain(levelInd(0));
  });

  it("pressing Enter again on the item now at the first level leaves the list", () => {
    const state = openState(item("One") + emptyItem(1), threeLevelNumbering());
    const once = runCommand(
      select(state, insideEmpty(state, 1)),
      docxKeymap.Enter
    );
    const twice = runCommand(
      select(once, insideEmpty(once, 1)),
      docxKeymap.Enter
    );
    const abandoned = paragraphAt(twice.doc, 1);

    expect(listRef(abandoned)).toBeNull();
    expect(pPrOf(abandoned)).toBe("");
    expect(twice.selection.$from.parentOffset).toBe(0);
  });

  it("an item that has text splits the way any paragraph does", () => {
    const state = openState(item("One"), threeLevelNumbering());
    const next = runCommand(
      select(state, posOfText(state.doc, "One") + 1),
      docxKeymap.Enter
    );
    const opened = paragraphAt(next.doc, 1);

    expect(next.doc.childCount).toBe(2);
    expect(listRef(opened)).toEqual({ numId: 3, ilvl: 0 });
    expect(pPrOf(opened)).toContain(levelInd(0));
  });

  it("an empty paragraph that is not a list splits the way any paragraph does", () => {
    const state = openState(paragraph("Body text") + "<w:p/>");
    const next = runCommand(
      select(state, insideEmpty(state, 1)),
      docxKeymap.Enter
    );
    expect(next.doc.childCount).toBe(3);
  });

  it("the abandoned item carries neither the list nor an indentation into the exported xml", () => {
    const { doc, session } = importDocx(
      makeNumberedDocx(item("One") + emptyItem())
    );
    const state = editorStateForSession({ doc, session });
    const next = runCommand(
      select(state, insideEmpty(state, 1)),
      docxKeymap.Enter
    );
    const xml = documentXmlOf(next.doc, session);

    // The item that kept its text keeps its list slot and the indentation the document wrote
    expect(xml).toContain('<w:numId w:val="3"/>');
    expect(xml).toContain('<w:ind w:left="720" w:hanging="360"/>');
    // The abandoned one goes out as a bare paragraph, carrying no formatting at all
    expect(xml).toContain("</w:p><w:p></w:p></w:body>");
  });
});

describe("a document with no numbering.xml", () => {
  /** A single-page document with nowhere to write a new list definition */
  function openWithoutNumberingPart(body: string): EditorState {
    const { doc, session } = importDocx(makeDocx(body));
    return editorStateForSession({ doc, session });
  }

  it("does not start a new list", () => {
    const state = openWithoutNumberingPart(paragraph("Body"));
    const at = select(state, posOfText(state.doc, "Body"));

    expect(toggleNumberedList(at, () => undefined)).toBe(false);
    expect(toggleBulletList(at, () => undefined)).toBe(false);
  });

  /** A list the document was already using needs no new definition written, so it can be edited as is */
  it("edits a list that already existed just as before", () => {
    const state = openWithoutNumberingPart(
      paragraph("Entry", listPPr(3, 0, '<w:ind w:left="720" w:hanging="360"/>'))
    );
    const at = select(state, posOfText(state.doc, "Entry"));

    expect(
      listRef(paragraphAt(runCommand(at, increaseListLevel).doc, 0))?.ilvl
    ).toBe(1);
    expect(
      listRef(paragraphAt(runCommand(at, removeFromList).doc, 0))
    ).toBeNull();
    // Nothing here defines list 3, so neither button is pressed and either would start a list,
    // which is what this document has nowhere to write. `removeFromList` above is the way out
    expect(toggleBulletList(at, () => undefined)).toBe(false);
    expect(toggleNumberedList(at, () => undefined)).toBe(false);
  });
});
