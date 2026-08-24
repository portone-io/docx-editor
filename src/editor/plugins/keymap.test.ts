// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { serializeParagraph } from "../../docx/serializeParagraph";
import {
  isBoldActive,
  isItalicActive,
  isStrikeActive,
  isUnderlineActive,
} from "../commands/formattingCommands";
import { documentHasLocked, lockSelection } from "../commands/lockCommands";
import { createEditorState } from "../createEditor";
import { docxKeymap } from "./keymap";
import { isLinkPanelOpen } from "./linkPanel";

/** A document of a single paragraph containing nothing but text */
function opened(text: string): EditorState {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  return createEditorState(importDocx(makeDocx(body)).doc);
}

function tableFollowedByParagraph(after = "", following = ""): EditorState {
  const cell = "<w:tc><w:p/></w:tc>";
  const body = `<w:tbl><w:tr>${cell}</w:tr></w:tbl><w:p>${after}</w:p>${following}`;
  return createEditorState(importDocx(makeDocx(body)).doc);
}

describe("Backspace", () => {
  it("does nothing of its own with the caret in the middle of a line", () => {
    expect(docxKeymap.Backspace(select(opened("Body text"), 5, 5))).toBe(false);
  });

  it("does nothing of its own with text selected", () => {
    expect(docxKeymap.Backspace(select(opened("Body text"), 1, 5))).toBe(false);
  });

  it("does nothing of its own right after text was typed", () => {
    const state = select(opened("Body text"), 5, 5);
    const typed = state.apply(state.tr.insertText("x"));
    expect(docxKeymap.Backspace(typed)).toBe(false);
  });

  it("keeps the final empty paragraph after a table", () => {
    const state = tableFollowedByParagraph();
    const caret = select(state, state.doc.content.size - 1);
    let dispatched = false;

    expect(
      docxKeymap.Backspace(caret, () => {
        dispatched = true;
      })
    ).toBe(true);
    expect(dispatched).toBe(false);
    expect(caret.doc.lastChild?.type.name).toBe("paragraph");
  });

  it("keeps an empty paragraph between two tables", () => {
    const followingTable = "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>";
    const state = tableFollowedByParagraph("", followingTable);
    const paragraphStart = state.doc.child(0).nodeSize + 1;

    expect(docxKeymap.Backspace(select(state, paragraphStart))).toBe(true);
    expect(state.doc.child(1).type.name).toBe("paragraph");
  });

  it("leaves a non-empty paragraph after a table to the base keymap", () => {
    const state = tableFollowedByParagraph(
      '<w:r><w:t xml:space="preserve">after</w:t></w:r>'
    );
    expect(
      docxKeymap.Backspace(select(state, state.doc.content.size - 1))
    ).toBe(false);
  });
});

describe("formatting shortcuts", () => {
  const shortcuts: ReadonlyArray<
    [key: string, isActive: (state: EditorState) => boolean]
  > = [
    ["Mod-b", isBoldActive],
    ["Mod-i", isItalicActive],
    ["Mod-u", isUnderlineActive],
    ["Mod-Shift-x", isStrikeActive],
  ];

  it.each(shortcuts)(
    "%s turns the formatting on and off again",
    (key, isActive) => {
      const selected = select(opened("abc"), 1, 3);
      const on = runCommand(selected, docxKeymap[key]);
      expect(isActive(on)).toBe(true);
      expect(isActive(runCommand(on, docxKeymap[key]))).toBe(false);
    }
  );
});

/**
 * The lock guard refuses the step a replayed lock is made of, so the binding has to be the
 * editor's own undo, which carries the pass for it (`editor/historyCommands`).
 */
describe("the history shortcuts", () => {
  it("take a lock back and shut it again", () => {
    const locked = runCommand(select(opened("Body text"), 1, 5), lockSelection);
    expect(documentHasLocked(locked.doc)).toBe(true);

    const back = runCommand(locked, docxKeymap["Mod-z"]);
    expect(documentHasLocked(back.doc)).toBe(false);
    expect(
      documentHasLocked(runCommand(back, docxKeymap["Shift-Mod-z"]).doc)
    ).toBe(true);
  });
});

describe("the break shortcuts", () => {
  const brOf = (state: EditorState) =>
    serializeParagraph(state.doc.child(0)).match(/<w:br[^>]*\/>/g) ?? [];

  it("Shift-Enter breaks the line without starting a page", () => {
    const next = runCommand(
      select(opened("abcd"), 3),
      docxKeymap["Shift-Enter"]
    );
    expect(brOf(next)).toEqual(["<w:br/>"]);
  });

  it("Mod-Enter starts the next page", () => {
    const next = runCommand(select(opened("abcd"), 3), docxKeymap["Mod-Enter"]);
    expect(brOf(next)).toEqual(['<w:br w:type="page"/>']);
  });

  // Answering false is what leaves it alone: nothing is dispatched, so there is no break to find
  it.each(["Shift-Enter", "Mod-Enter"])(
    "%s leaves a locked stretch alone and says so",
    (key) => {
      const locked = runCommand(select(opened("abcd"), 1, 5), lockSelection);

      expect(docxKeymap[key](select(locked, 3))).toBe(false);
    }
  );
});

describe("Tab", () => {
  it("inserts a document tab in an ordinary paragraph", () => {
    const next = runCommand(select(opened("abcd"), 3), docxKeymap.Tab);
    expect(next.doc.child(0).textContent).toBe("ab\tcd");
    expect(serializeParagraph(next.doc.child(0))).toContain("<w:tab/>");
  });

  it("is consumed without inserting a tab when a list cannot move deeper", () => {
    const body =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="8"/>' +
      '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>';
    const state = createEditorState(importDocx(makeDocx(body)).doc);
    let dispatched = false;

    expect(
      docxKeymap.Tab(select(state, 2), () => {
        dispatched = true;
      })
    ).toBe(true);
    expect(dispatched).toBe(false);
    expect(state.doc.textContent).toBe("item");
  });
});

describe("the link key", () => {
  it("opens the link panel over selected text", () => {
    const opening = runCommand(
      select(opened("abc"), 1, 4),
      docxKeymap["Mod-k"]
    );
    expect(isLinkPanelOpen(opening)).toBe(true);
  });

  // Answering false leaves the key to the browser, which has one of its own on it
  it("does nothing where there is nothing to link", () => {
    expect(docxKeymap["Mod-k"](select(opened("abc"), 2))).toBe(false);
  });

  it("leaves a locked stretch alone and says so", () => {
    const locked = runCommand(select(opened("abcd"), 1, 5), lockSelection);
    expect(docxKeymap["Mod-k"](select(locked, 1, 5))).toBe(false);
  });
});
