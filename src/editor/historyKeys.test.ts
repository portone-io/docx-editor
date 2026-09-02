// @vitest-environment jsdom
import { closeHistory } from "prosemirror-history";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { rangeOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import type { EditingProtection } from "../schema/protection";
import { addComment } from "./commands/commentCommands";
import { createEditorState, createEditorView } from "./createEditor";
import { documentDefaults } from "./documentStyles";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const BODY = `<w:p>${run("Alpha ")}${run("beta")}</w:p>`;

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function openView(protection: EditingProtection): EditorView {
  const mount = document.body.appendChild(document.createElement("div"));
  const state = createEditorState(importDocx(makeDocx(BODY)).doc, {
    protection,
    author: { id: "me", name: "Me" },
  });
  const view = createEditorView({
    mount,
    state,
    defaults: documentDefaults(state),
    readOnly: false,
    onStateChange: () => {},
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

/** prosemirror-keymap reads `Mod-` as Cmd on a Mac and as Ctrl everywhere else */
const mac = /Mac|iP(hone|[oa]d)/.test(navigator.platform);

function pressUndo(view: EditorView): KeyboardEvent {
  const event = new KeyboardEvent("keydown", {
    key: "z",
    keyCode: 90,
    ctrlKey: !mac,
    metaKey: mac,
    bubbles: true,
    cancelable: true,
  });
  view.dom.dispatchEvent(event);
  return event;
}

function commentCount(view: EditorView): number {
  let count = 0;
  view.state.doc.descendants((node) => {
    if (node.type.name === "commentReference") count += 1;
  });
  return count;
}

describe("the undo key", () => {
  it("takes a press while the sheet is not editable", () => {
    const view = openView("comments");
    expect(view.editable).toBe(false);

    const { from, to } = rangeOfText(view.state.doc, "beta");
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
    );
    addComment({ text: "note", author: "Me", authorId: "me" })(
      view.state,
      (tr) => view.dispatch(tr)
    );
    expect(commentCount(view)).toBe(1);

    const event = pressUndo(view);
    expect(commentCount(view)).toBe(0);
    expect(event.defaultPrevented).toBe(true);
  });

  it("still runs through the keymap while the sheet is editable, one edit per press", () => {
    const view = openView("none");
    expect(view.editable).toBe(true);

    view.dispatch(view.state.tr.insertText("one", 1));
    view.dispatch(closeHistory(view.state.tr).insertText("two", 1));
    expect(view.state.doc.textContent).toBe("twooneAlpha beta");

    // A second handler running the same binding would take both edits back at once
    expect(pressUndo(view).defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("oneAlpha beta");

    pressUndo(view);
    expect(view.state.doc.textContent).toBe("Alpha beta");
  });
});
