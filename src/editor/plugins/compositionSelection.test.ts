// @vitest-environment jsdom
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { openComposition } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { createEditorState, createEditorView } from "../createEditor";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const PARAGRAPH = `<w:p>${run("Source text")}</w:p>`;

/** Plain text, a locked control, then plain text again */
const WITH_LOCK =
  `<w:p>${run("pre")}<w:sdt>${LOCKED_PR}<w:sdtContent>${run("locked")}` +
  `</w:sdtContent></w:sdt>${run("post")}</w:p>`;

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function openEditor(body: string): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const { doc } = importDocx(makeDocx(body));
  const view = createEditorView({
    mount,
    state: createEditorState(doc),
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

function select(view: EditorView, from: number, to: number): void {
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
  );
}

describe("a composition opening over selected text", () => {
  it("takes the selected text away before the browser composes", () => {
    const view = openEditor(PARAGRAPH);
    select(view, 1, 7);

    openComposition(view);

    expect(view.state.doc.textContent).toBe(" text");
    expect(view.state.selection.empty).toBe(true);
    expect(view.state.selection.from).toBe(1);
  });

  it("leaves a caret alone, which the browser composes at as it stands", () => {
    const view = openEditor(PARAGRAPH);
    select(view, 4, 4);

    openComposition(view);

    expect(view.state.doc.textContent).toBe("Source text");
    expect(view.state.selection.from).toBe(4);
  });

  it("leaves the document as it stands where a guard refuses the deletion", () => {
    const view = openEditor(WITH_LOCK);
    // Across the locked control, which no edit may replace
    select(view, 1, 14);

    openComposition(view);

    expect(view.state.doc.textContent).toBe("prelockedpost");
    expect(view.state.selection.from).toBe(1);
  });
});
