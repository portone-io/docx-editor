// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
} from "../../__testing__/docx";
import { posOfText, runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toParagraphFormat } from "../../model/format";
import { parseNumbering } from "../../numbering/parseNumbering";
import { createEditorState } from "../createEditor";
import {
  canDecreaseIndent,
  canIncreaseIndent,
  decreaseIndent,
  increaseIndent,
} from "./indentCommands";
import { toggleNumberedList } from "./listCommands";

function paragraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

interface Opened {
  state: EditorState;
  session: SessionStore;
}

function opened(body: string): Opened {
  const { doc, session } = importDocx(makeDocx(body));
  return { state: createEditorState(doc, { styles: session.styles }), session };
}

/** A document that can start a list, so a list paragraph can be reached from the commands */
function openedNumbered(body: string): Opened {
  const { doc, session } = importDocx(makeNumberedDocx(body));
  return {
    state: createEditorState(doc, {
      numbering: parseNumbering(session.numberingXml),
      styles: session.styles,
    }),
    session,
  };
}

/** A state with the caret placed on the first character of the given text */
function at(state: EditorState, text: string): EditorState {
  return select(state, posOfText(state.doc, text));
}

/** Runs the commands one after another, each with the caret in that paragraph */
function runAll(
  state: EditorState,
  text: string,
  commands: readonly Command[]
): EditorState {
  return commands.reduce(
    (current, command) => runCommand(at(current, text), command),
    state
  );
}

function pPrOf(doc: PMNode, index: number): string {
  const pPr: unknown = doc.child(index).attrs.pPr;
  return typeof pPr === "string" ? pPr : "";
}

function indentStartPtOf(doc: PMNode, index: number): number | undefined {
  return toParagraphFormat(doc.child(index).attrs.format)?.indentStartPt;
}

describe("indenting a plain paragraph", () => {
  const TEXT = "Body text";

  it("moves it half an inch to the right", () => {
    const { state, session } = opened(paragraph(TEXT));
    const indented = runCommand(at(state, TEXT), increaseIndent);

    expect(indentStartPtOf(indented.doc, 0)).toBe(36);
    expect(documentXmlOf(indented.doc, session)).toContain(
      '<w:pPr><w:ind w:left="720"/></w:pPr>'
    );
  });

  it("each further step adds another half inch", () => {
    const { state } = opened(paragraph(TEXT));
    const twice = runAll(state, TEXT, [increaseIndent, increaseIndent]);

    expect(pPrOf(twice.doc, 0)).toContain('<w:ind w:left="1440"/>');
  });

  it("an outdent takes one step back off an indent the document wrote", () => {
    const pPr = '<w:pPr><w:ind w:left="1440"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    const out = runCommand(at(state, TEXT), decreaseIndent);

    expect(pPrOf(out.doc, 0)).toBe('<w:pPr><w:ind w:left="720"/></w:pPr>');
  });

  it("stops at the left margin instead of running past it", () => {
    const pPr = '<w:pPr><w:ind w:left="400"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    const out = runCommand(at(state, TEXT), decreaseIndent);

    expect(indentStartPtOf(out.doc, 0)).toBeUndefined();
    // A paragraph sitting at the margin has nothing left to give up
    expect(canDecreaseIndent(at(out, TEXT))).toBe(false);
    expect(decreaseIndent(at(out, TEXT))).toBe(false);
    expect(canIncreaseIndent(at(out, TEXT))).toBe(true);
  });

  it("comes back to the indent it started with", () => {
    const original = '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>';
    const { state, session } = opened(paragraph(TEXT, original));
    const back = runAll(state, TEXT, [increaseIndent, decreaseIndent]);

    expect(pPrOf(back.doc, 0)).toBe(original);
    expect(documentXmlOf(back.doc, session)).toContain(original);
  });

  it("a paragraph that had no indent is left with none again", () => {
    const { state, session } = opened(paragraph(TEXT));
    const back = runAll(state, TEXT, [increaseIndent, decreaseIndent]);

    expect(pPrOf(back.doc, 0)).toBe("");
    expect(documentXmlOf(back.doc, session)).not.toContain("<w:ind");
  });

  it("the hanging indent and the rest of the formatting stay as they were", () => {
    const pPr =
      '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>' +
      '<w:ind w:left="720" w:right="200" w:hanging="360"/>' +
      '<w:jc w:val="both"/></w:pPr>';
    const { state, session } = opened(paragraph(TEXT, pPr));
    const indented = runCommand(at(state, TEXT), increaseIndent);

    expect(documentXmlOf(indented.doc, session)).toContain(
      '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>' +
        '<w:ind w:left="1440" w:right="200" w:hanging="360"/>' +
        '<w:jc w:val="both"/></w:pPr>'
    );
  });

  it("a first-line indent survives the step", () => {
    const pPr = '<w:pPr><w:ind w:left="400" w:firstLine="200"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    const indented = runCommand(at(state, TEXT), increaseIndent);

    expect(pPrOf(indented.doc, 0)).toBe(
      '<w:pPr><w:ind w:left="1120" w:firstLine="200"/></w:pPr>'
    );
  });

  it("steps on from the indent a style passes down", () => {
    const { doc, session } = importDocx(
      makeStyledDocx(
        paragraph(TEXT, '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>'),
        '<w:style w:styleId="Quote">' +
          '<w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
      )
    );
    const state = createEditorState(doc, { styles: session.styles });
    const indented = runCommand(at(state, TEXT), increaseIndent);

    expect(pPrOf(indented.doc, 0)).toBe(
      '<w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="1440"/></w:pPr>'
    );
  });

  it("every selected paragraph takes its own step", () => {
    const { state } = opened(
      paragraph("First", '<w:pPr><w:ind w:left="720"/></w:pPr>') +
        paragraph("Second")
    );
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    const indented = runCommand(both, increaseIndent);

    expect(indentStartPtOf(indented.doc, 0)).toBe(72);
    expect(indentStartPtOf(indented.doc, 1)).toBe(36);
  });
});

describe("indenting a list paragraph", () => {
  const TEXT = "Item";

  /** The list slot the paragraph sits in */
  function listRefOf(doc: PMNode, index: number) {
    return toParagraphFormat(doc.child(index).attrs.format)?.numbering ?? null;
  }

  it("moves the item a level instead of its indent", () => {
    const { state } = openedNumbered(paragraph(TEXT));
    const deeper = runAll(state, TEXT, [toggleNumberedList, increaseIndent]);

    expect(listRefOf(deeper.doc, 0)?.ilvl).toBe(1);
    // The level's own indentation is what moves the item, so the paragraph records none
    expect(pPrOf(deeper.doc, 0)).not.toContain("<w:ind");
  });

  it("moves it back a level", () => {
    const { state } = openedNumbered(paragraph(TEXT));
    const back = runAll(state, TEXT, [
      toggleNumberedList,
      increaseIndent,
      decreaseIndent,
    ]);

    expect(listRefOf(back.doc, 0)?.ilvl).toBe(0);
    expect(pPrOf(back.doc, 0)).not.toContain("<w:ind");
  });

  it("the first level is as far left as an item goes", () => {
    const { state } = openedNumbered(paragraph(TEXT));
    const listed = runCommand(at(state, TEXT), toggleNumberedList);
    const spot = at(listed, TEXT);

    expect(canDecreaseIndent(spot)).toBe(false);
    expect(canIncreaseIndent(spot)).toBe(true);
  });

  it("a list item and a plain paragraph selected together each follow their own rule", () => {
    const { state } = openedNumbered(paragraph("Item") + paragraph("Body"));
    const listed = runCommand(at(state, "Item"), toggleNumberedList);
    const both = select(
      listed,
      posOfText(listed.doc, "Item"),
      posOfText(listed.doc, "Body")
    );
    const indented = runCommand(both, increaseIndent);

    expect(listRefOf(indented.doc, 0)?.ilvl).toBe(1);
    expect(indentStartPtOf(indented.doc, 0)).toBeUndefined();
    expect(indentStartPtOf(indented.doc, 1)).toBe(36);
  });
});
