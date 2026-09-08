// @vitest-environment jsdom
import { Fragment, Slice } from "prosemirror-model";
import { NodeSelection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { toRunFormat } from "../model/format";
import { docxSchema } from "../schema";
import { createEditorState, createEditorView } from "./createEditor";

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

const SOURCE = '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';

/** An editor with the caret placed at the start of its first paragraph */
function openEditor(body: string = SOURCE): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const { doc } = importDocx(makeDocx(body));
  const view = createEditorView({
    mount,
    state: createEditorState(doc),
    onStateChange: () => undefined,
  });
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))
  );
  // jsdom does not measure text, so a position cannot be found from coordinates. Pin the drop position to the start of the paragraph
  view.posAtCoords = () => ({ pos: 1, inside: 0 });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

/** Fakes a single drop. jsdom has neither DragEvent nor DataTransfer */
function drop(view: EditorView, text: string, html: string): boolean {
  const event = new Event("drop", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "dataTransfer", {
    value: {
      types: ["text/plain", "text/html"],
      files: [],
      getData: (type: string) => (type === "text/html" ? html : text),
    },
  });
  view.dom.dispatchEvent(event);
  return event.defaultPrevented;
}

/** The fragments the document is opened with and only preserves, which no edit may lose */
function markerCount(view: EditorView): number {
  let markers = 0;
  view.state.doc.descendants((node) => {
    if (node.attrs.guarded === true) markers += 1;
    return true;
  });
  return markers;
}

function hasBoldRun(view: EditorView): boolean {
  let bold = false;
  view.state.doc.descendants((node) => {
    for (const mark of node.marks) {
      if (toRunFormat(mark.attrs.format)?.bold) bold = true;
    }
    return true;
  });
  return bold;
}

describe("drag and drop", () => {
  it("drops the formatting and takes only the text when the content comes from outside", () => {
    const view = openEditor();
    const html =
      '<span class="docx-editor-run" data-fmt=\'{"bold":true}\'>intruder</span>';
    expect(drop(view, "intruder", html)).toBe(true);

    expect(view.state.doc.textContent).toBe("intrudersource");
    expect(hasBoldRun(view)).toBe(false);
  });

  it("does not split paragraphs on line breaks coming from outside", () => {
    const view = openEditor();
    drop(view, "a\r\nb", "<p>a</p><p>b</p>");

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.textContent).toBe("absource");
  });

  it("leaves ProseMirror to do what it always does while dragging inside the editor", () => {
    const view = openEditor();
    view.dragging = {
      slice: new Slice(Fragment.from(docxSchema.text("moved")), 0, 0),
      move: false,
    };
    // The slice being dragged must be what gets inserted, not the clipboard text
    drop(view, "clipboard", "clipboard");

    expect(view.state.doc.textContent).toBe("movedsource");
  });

  it.each([
    { move: true, kept: "kept" },
    { move: false, kept: null },
  ])(
    "keeps the source identity of a block only while the drag moves it (move: $move)",
    ({ move, kept }) => {
      const view = openEditor();
      view.dragging = {
        slice: new Slice(
          Fragment.from(
            docxSchema.nodes.paragraph.create({ srcId: "kept" }, [
              docxSchema.text("dragged"),
            ])
          ),
          0,
          0
        ),
        move,
      };
      drop(view, "clipboard", "clipboard");

      // A copy is written from the model; only the block that moved is still the block it was
      expect(view.state.doc.firstChild?.attrs.srcId).toBe(kept);
    }
  );
  it("carries the markers a moved block anchors along with it", () => {
    const view = openEditor(
      SOURCE +
        '<w:p><w:bookmarkStart w:id="1" w:name="mark"/>' +
        '<w:r><w:t xml:space="preserve">anchored</w:t></w:r>' +
        '<w:bookmarkEnd w:id="1"/></w:p>'
    );
    const anchored = view.state.doc.firstChild?.nodeSize ?? 0;
    view.dispatch(
      view.state.tr.setSelection(NodeSelection.create(view.state.doc, anchored))
    );
    view.dragging = { slice: view.state.selection.content(), move: true };
    drop(view, "anchored", "anchored");

    // Taking the markers off would leave a document the file cannot be written back from, which
    // the preserved guard refuses whole, so the drag would silently land nothing at all
    expect(view.state.doc.lastChild?.textContent).toBe("source");
    expect(markerCount(view)).toBe(2);
  });
});
