// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeNotesDocx, NOTE_BODY } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { docxSchema } from "../../schema";
import { storyKey } from "../../schema/stories";
import { createEditorView, editorStateForSession } from "../createEditor";
import { openNoteCommand, requestedNote } from "./noteNavigation";

let opened: EditorView | null = null;

afterEach(() => {
  opened?.destroy();
  opened = null;
});

function mainView(body: string = NOTE_BODY): EditorView {
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(importDocx(makeNotesDocx(body))),
    onStateChange: () => {},
  });
  opened = view;
  return view;
}

/** Where the reference to a note of this kind stands, and the node itself */
function referenceOf(doc: PMNode, kind: string): { pos: number; node: PMNode } {
  let found: { pos: number; node: PMNode } | null = null;
  doc.descendants((node, pos) => {
    if (
      found === null &&
      node.type === docxSchema.nodes.noteReference &&
      node.attrs.kind === kind
    ) {
      found = { pos, node };
    }
    return found === null;
  });
  if (found === null) throw new Error(`no ${kind} reference`);
  return found;
}

/** A press on the number, which is what `handleClickOn` is asked about */
function clickOn(view: EditorView, kind: string): boolean {
  const { pos, node } = referenceOf(view.state.doc, kind);
  return (
    view.someProp("handleClickOn", (run) =>
      run(view, pos, node, pos, new MouseEvent("mousedown"), true)
    ) === true
  );
}

describe("getting into a note from its number", () => {
  it("asks for the footnote whose number was clicked", () => {
    const view = mainView();
    expect(requestedNote(view.state)).toBeNull();

    expect(clickOn(view, "footnote")).toBe(true);

    expect(requestedNote(view.state)?.key).toBe(storyKey("footnote", "2"));
    // The caret is left where Escape hands it back to, just after the reference
    const { pos, node } = referenceOf(view.state.doc, "footnote");
    expect(view.state.selection.from).toBe(pos + node.nodeSize);
  });

  it("asks for nothing when an endnote number is clicked", () => {
    const view = mainView();

    expect(clickOn(view, "endnote")).toBe(false);

    expect(requestedNote(view.state)).toBeNull();
  });

  it("asks again for the note already open, so a second press still reaches the surface", () => {
    const view = mainView();
    clickOn(view, "footnote");
    const first = requestedNote(view.state);

    clickOn(view, "footnote");

    const second = requestedNote(view.state);
    expect(second?.key).toBe(first?.key);
    expect(second?.nonce).toBeGreaterThan(first?.nonce ?? 0);
  });

  it("opens no note the document refers to nothing of", () => {
    const view = mainView();

    expect(openNoteCommand("footnote", "404")(view.state)).toBe(false);
    expect(requestedNote(view.state)).toBeNull();
  });
});
