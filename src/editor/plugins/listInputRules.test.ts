// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toParagraphFormat } from "../../model/format";
import { type Numbering, parseNumbering } from "../../numbering/parseNumbering";
import { activeListKind, toggleNumberedList } from "../commands/listCommands";
import { createEditorState, createEditorView } from "../createEditor";
import { docxKeymap } from "./keymap";
import { paragraphMarkers } from "./numberingDecorations";

/** A paragraph with nothing typed into it yet */
const EMPTY_PARAGRAPH = "<w:p/>";

function bodyParagraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function listPPr(numId: number, ilvl: number): string {
  return (
    `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/>` +
    `<w:numId w:val="${numId}"/></w:numPr></w:pPr>`
  );
}

interface Opened {
  view: EditorView;
  session: SessionStore;
  numbering: Numbering;
}

const mounted: { view: EditorView | null } = { view: null };

afterEach(() => {
  mounted.view?.destroy();
  mounted.view = null;
});

function openState(bytes: Uint8Array): {
  state: EditorState;
  session: SessionStore;
} {
  const { doc, session } = importDocx(bytes);
  return {
    state: createEditorState(doc, {
      numbering: parseNumbering(session.numberingXml),
      styles: session.styles,
      defaults: session.defaults,
      canStartNewList: session.numberingPartPath !== null,
    }),
    session,
  };
}

/** Opens the body in a mounted editor. Without `numbering` the document has no numbering.xml, so no list can be started in it */
function open(body: string, options: { numbering?: boolean } = {}): Opened {
  const withNumbering = options.numbering !== false;
  const { state, session } = openState(
    withNumbering ? makeNumberedDocx(body) : makeDocx(body)
  );
  const view = createEditorView({
    mount: document.createElement("div"),
    state,
    defaults: session.defaults,
    readOnly: false,
    onStateChange: () => {},
  });
  mounted.view = view;
  return { view, session, numbering: parseNumbering(session.numberingXml) };
}

/** Puts the caret at a character offset inside the paragraph at `index` */
function caretAt(view: EditorView, index: number, offset: number): void {
  let start = 1;
  for (let before = 0; before < index; before += 1) {
    start += view.state.doc.child(before).nodeSize;
  }
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.create(view.state.doc, start + offset)
    )
  );
}

/**
 * Offers the text to the input rules the way the view does when a key is pressed, without
 * inserting it. Answers whether a rule took it
 */
function offerToRules(
  view: EditorView,
  from: number,
  to: number,
  text: string
): boolean {
  const insert = () => view.state.tr.insertText(text, from, to);
  return (
    view.someProp("handleTextInput", (handle) =>
      handle(view, from, to, text, insert)
    ) === true
  );
}

/**
 * Types one character at a time. Each is offered to the input rules the way the browser does,
 * and inserted here only when no rule took it
 */
function type(view: EditorView, text: string): void {
  for (const character of text) {
    const { from, to } = view.state.selection;
    if (offerToRules(view, from, to, character)) continue;
    view.dispatch(view.state.tr.insertText(character, from, to));
  }
}

function press(view: EditorView, command: Command): boolean {
  return command(view.state, (tr) => view.dispatch(tr), view);
}

function paragraphAt(view: EditorView, index = 0): PMNode {
  return view.state.doc.child(index);
}

function listRef(node: PMNode) {
  return toParagraphFormat(node.attrs.format)?.numbering ?? null;
}

function pPrOf(node: PMNode): string {
  const pPr: unknown = node.attrs.pPr;
  return typeof pPr === "string" ? pPr : "";
}

/** The numbers drawn on screen, in document order */
function markerTexts({ view, numbering }: Opened): string[] {
  return paragraphMarkers(view.state.doc, numbering).map(
    (marker) => marker.text
  );
}

/** The same body turned into a list by the toolbar button, with the caret in the first paragraph */
function listedByButton(body: string): {
  doc: PMNode;
  session: SessionStore;
} {
  const { state, session } = openState(makeNumberedDocx(body));
  const selected = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 1))
  );
  let listed = selected;
  expect(
    toggleNumberedList(selected, (tr) => {
      listed = selected.apply(tr);
    })
  ).toBe(true);
  return { doc: listed.doc, session };
}

describe('typing "1. " at the start of a line', () => {
  it("starts a numbered list and takes the prefix away", () => {
    const opened = open(EMPTY_PARAGRAPH);
    caretAt(opened.view, 0, 0);
    type(opened.view, "1. ");
    const item = paragraphAt(opened.view);

    expect(item.textContent).toBe("");
    expect(listRef(item)).toEqual({ numId: 2, ilvl: 0 });
    expect(markerTexts(opened)).toEqual(["1."]);
    // The level's indentation is drawn on screen, so the paragraph records none of its own
    expect(pPrOf(item)).not.toContain("<w:ind");
  });

  it("leaves the caret at the start of the item", () => {
    const { view } = open(EMPTY_PARAGRAPH);
    caretAt(view, 0, 0);
    type(view, "1. ");
    expect(view.state.selection.$from.parentOffset).toBe(0);
  });

  it("converts a line that already has text, in front of that text", () => {
    const opened = open(bodyParagraph("Body text"));
    caretAt(opened.view, 0, 0);
    type(opened.view, "1. ");

    expect(paragraphAt(opened.view).textContent).toBe("Body text");
    expect(markerTexts(opened)).toEqual(["1."]);
  });

  it("leaves the document the list button leaves", () => {
    const { view } = open(bodyParagraph("Body text"));
    caretAt(view, 0, 0);
    type(view, "1. ");

    expect(
      view.state.doc.eq(listedByButton(bodyParagraph("Body text")).doc)
    ).toBe(true);
  });

  it("exports the same xml a list made with the button does", () => {
    const body = bodyParagraph("Body text");
    const { view, session } = open(body);
    caretAt(view, 0, 0);
    type(view, "1. ");
    const byButton = listedByButton(body);
    const xml = documentXmlOf(view.state.doc, session);

    expect(xml).toBe(documentXmlOf(byButton.doc, byButton.session));
    expect(xml).toContain("<w:numPr>");
    expect(xml).not.toContain("<w:ind");
    expect(xml).not.toContain("1. ");
  });

  /**
   * A list started this way counts from one, so any other number would come back renumbered and
   * the number that was typed would be lost. Word writes the number into the list definition
   * instead; until this editor does the same, only `1.` converts.
   */
  it("no other number converts", () => {
    const opened = open(EMPTY_PARAGRAPH);
    caretAt(opened.view, 0, 0);
    type(opened.view, "7. ");

    expect(paragraphAt(opened.view).textContent).toBe("7. ");
    expect(listRef(paragraphAt(opened.view))).toBeNull();
  });
});

describe("typing a bullet prefix", () => {
  it.each(["- ", "* "])("%j starts a bullet list", (prefix) => {
    const opened = open(EMPTY_PARAGRAPH);
    caretAt(opened.view, 0, 0);
    type(opened.view, prefix);

    expect(paragraphAt(opened.view).textContent).toBe("");
    expect(activeListKind(opened.view.state)).toBe("bullet");
    expect(markerTexts(opened)).toEqual(["●"]);
  });
});

describe("taking the conversion back", () => {
  function converted(): EditorView {
    const { view } = open(EMPTY_PARAGRAPH);
    caretAt(view, 0, 0);
    type(view, "1. ");
    return view;
  }

  it("Backspace puts the typed prefix back as text", () => {
    const view = converted();
    expect(press(view, docxKeymap.Backspace)).toBe(true);
    const item = paragraphAt(view);

    expect(item.textContent).toBe("1. ");
    expect(listRef(item)).toBeNull();
    expect(pPrOf(item)).toBe("");
  });

  it("a second Backspace is left to the base keymap again", () => {
    const view = converted();
    press(view, docxKeymap.Backspace);
    expect(docxKeymap.Backspace(view.state, undefined)).toBe(false);
  });

  it("Backspace does nothing of its own where no rule has just run", () => {
    const { view } = open(bodyParagraph("Body text"));
    caretAt(view, 0, 4);
    expect(docxKeymap.Backspace(view.state, undefined)).toBe(false);

    type(view, "x");
    expect(docxKeymap.Backspace(view.state, undefined)).toBe(false);
  });

  /**
   * The prefix and the conversion it caused are one group in the history, so a single undo takes
   * the line all the way back to what it was and never leaves it half converted
   */
  it("one undo takes the whole line back, and one redo brings the list back", () => {
    const view = converted();
    expect(press(view, docxKeymap["Mod-z"])).toBe(true);
    expect(listRef(paragraphAt(view))).toBeNull();
    expect(paragraphAt(view).textContent).toBe("");

    expect(press(view, docxKeymap["Mod-y"])).toBe(true);
    expect(listRef(paragraphAt(view))).toEqual({ numId: 2, ilvl: 0 });
  });
});

describe("the prefixes that are not read as a list", () => {
  it("nothing happens in a paragraph that is already an item", () => {
    const opened = open(`<w:p>${listPPr(1, 0)}</w:p>`);
    caretAt(opened.view, 0, 0);
    type(opened.view, "1. ");
    const item = paragraphAt(opened.view);

    expect(item.textContent).toBe("1. ");
    expect(listRef(item)).toEqual({ numId: 1, ilvl: 0 });
  });

  it("nothing happens in the middle of a line", () => {
    const { view } = open(bodyParagraph("Body"));
    caretAt(view, 0, 4);
    type(view, "1. ");

    expect(paragraphAt(view).textContent).toBe("Body1. ");
    expect(listRef(paragraphAt(view))).toBeNull();
  });

  it("nothing happens in a document that has no numbering.xml to define a list in", () => {
    const { view } = open(EMPTY_PARAGRAPH, { numbering: false });
    caretAt(view, 0, 0);
    type(view, "1. ");

    expect(paragraphAt(view).textContent).toBe("1. ");
    expect(listRef(paragraphAt(view))).toBeNull();
  });

  /**
   * A whole prefix can arrive as a single piece of input (dictation, or a phone keyboard
   * completing a word), and the place it arrives at need not be a line at all
   */
  it("a preserved block does not become one", () => {
    const { view } = open("<w:customXml/>");
    const before = view.state.doc;
    expect(before.child(0).type.name).toBe("docxRaw");

    expect(offerToRules(view, 0, 0, "1. ")).toBe(false);
    expect(view.state.doc).toBe(before);
  });
});
