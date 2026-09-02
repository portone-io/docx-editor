// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { keymap } from "prosemirror-keymap";
import {
  type Command,
  Plugin,
  PluginKey,
  TextSelection,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { act, type ReactNode, useEffect, useRef, useState } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx, readFixture } from "./__testing__/docx";
import { AUTHOR, EDITING } from "./__testing__/mode";
import { renderInto } from "./__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "./DocxEditor";
import { editingProtection } from "./editor/commands/index";
import { DocxImportError } from "./ooxml/errors";
import { editorClassNames } from "./styles/classNames";
import type { FontFallbacks } from "./styles/fontStack";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

const ONE_PARAGRAPH = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
);

/** One run that declares a font, so the fallbacks after that name can be read off the DOM */
const BATANG_PARAGRAPH = makeDocx(
  '<w:p><w:r><w:rPr><w:rFonts w:ascii="Batang"/></w:rPr>' +
    '<w:t xml:space="preserve">source</w:t></w:r></w:p>'
);

const OTHER_PARAGRAPH = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">other text</w:t></w:r></w:p>'
);

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

/** A paragraph to right click, and a table cell to right click, which open different menus */
const PARAGRAPH_AND_TABLE = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>' +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    `<w:tr>${cellXml("Cell")}</w:tr>` +
    "</w:tbl>"
);

const render = (element: ReactNode) => renderInto(host, element);

/** The page marks find their positions on an animation frame, which this waits out */
function nextFrame() {
  return act(async () => {
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => resolve());
    });
  });
}

/** A second of frames, which is the point at which a draw that never came is a failure */
const DRAWING_FRAMES = 60;

/** Waits out a blob read and the render it lands in */
function settled() {
  return act(async () => {
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 0);
    });
  });
}

/** The same bytes, copied into a buffer of their own, which is what a Blob part asks for */
function copyOf(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy;
}

const fileOf = (bytes: Uint8Array, name: string) =>
  new File([copyOf(bytes)], name);

/**
 * Takes frames until that many guides stand drawn.
 *
 * Each frame is taken inside its own `act`, because what the measurement leaves behind is React
 * state and the draw it turns into only lands once the surrounding `act` has closed.
 */
async function untilGuides(drawn: () => number, count: number): Promise<void> {
  for (let frame = 0; frame < DRAWING_FRAMES && drawn() !== count; frame += 1) {
    await nextFrame();
  }
}

/** A box to hold whatever the ref hands back */
function handleBox(): { current: DocxEditorHandle | null } {
  return { current: null };
}

function attached(box: { current: DocxEditorHandle | null }): DocxEditorHandle {
  const editor = box.current;
  if (!editor) throw new Error("the ref was not attached");
  return editor;
}

describe("DocxEditor", () => {
  it("renders the document and lets the parent read the ref right away in its own effect", () => {
    const exported = { byteLength: 0 };

    function Host() {
      const editorRef = useRef<DocxEditorHandle | null>(null);
      useEffect(() => {
        exported.byteLength = editorRef.current?.exportBytes().length ?? 0;
      }, []);
      return (
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={EDITING}
          ref={editorRef}
          renderImportError={() => null}
        />
      );
    }

    const unmount = render(<Host />);
    expect(host.textContent).toContain("source");
    expect(exported.byteLength).toBeGreaterThan(0);
    unmount();
  });

  it("exports the edited result", () => {
    const box = handleBox();
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        ref={box}
        renderImportError={() => null}
      />
    );

    const editor = box.current;
    if (!editor) throw new Error("the ref was not attached");
    act(() => {
      editor.view.dispatch(editor.view.state.tr.insertText("edited ", 1));
    });

    const documentXml = decode(
      unzipSync(editor.exportBytes())["word/document.xml"]
    );
    expect(documentXml).toContain("edited source");
    unmount();
  });

  it("shows the screen to draw instead for a document that could not be opened", () => {
    const rejected: { current: DocxImportError | null } = { current: null };
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH.slice(0, 40)}
        mode={EDITING}
        renderImportError={(error) => {
          rejected.current = error;
          return <p>could not open</p>;
        }}
      />
    );
    expect(host.textContent).toBe("could not open");
    expect(rejected.current).toBeInstanceOf(DocxImportError);
    unmount();
  });

  it("draws a panel of its own for a refused document when none was handed in", () => {
    const unmount = render(
      <DocxEditor document={ONE_PARAGRAPH.slice(0, 40)} mode={EDITING} />
    );

    const panel = host.querySelector(`.${editorClassNames.rejection}`);
    expect(panel).not.toBeNull();
    // The reason the refusal carries is what the panel says, not an empty box
    expect(panel?.textContent).toContain("not opened");
    expect(panel?.textContent).toContain("not a docx document");
    // Nothing of the editor itself is standing
    expect(host.querySelector(`.${editorClassNames.sheet}`)).toBeNull();
    unmount();
  });

  it("opens bytes handed in as an ArrayBuffer without waiting", () => {
    const box = handleBox();

    const unmount = render(
      <DocxEditor
        document={copyOf(ONE_PARAGRAPH).buffer}
        mode={EDITING}
        ref={box}
      />
    );

    // Synchronously after the first render, with no frame taken in between
    expect(box.current?.view.state.doc.textContent).toContain("source");
    unmount();
  });

  it("reads a file the consumer hands over and opens what it holds", async () => {
    const box = handleBox();
    const file = fileOf(ONE_PARAGRAPH, "handed-over.docx");

    const unmount = render(
      <DocxEditor document={file} mode={EDITING} ref={box} />
    );

    // The blob is read first, so the editor stands empty for that moment
    expect(box.current).toBeNull();
    await settled();
    expect(box.current?.view.state.doc.textContent).toContain("source");
    unmount();
  });

  it("stands the old document down the moment another file is handed in", async () => {
    const box = handleBox();
    // Every document the editor ever put on screen, in the order it opened them
    const opened: string[] = [];
    let hand = (_file: File) => {};

    function Swapping() {
      const [file, setFile] = useState(() =>
        fileOf(ONE_PARAGRAPH, "first.docx")
      );
      hand = setFile;
      return (
        <DocxEditor
          document={file}
          mode={EDITING}
          ref={box}
          onReady={(view) => opened.push(view.state.doc.textContent)}
        />
      );
    }

    const unmount = render(<Swapping />);
    await settled();
    expect(opened).toEqual(["source"]);

    act(() => hand(fileOf(OTHER_PARAGRAPH, "second.docx")));

    // While the new file is being read the handle is detached, so nothing can export
    // the document that was asked to be replaced
    expect(box.current).toBeNull();
    expect(host.querySelector(`.${editorClassNames.sheet}`)).toBeNull();

    await settled();
    expect(box.current?.view.state.doc.textContent).toContain("other text");
    expect(opened).toEqual(["source", "other text"]);
    unmount();
  });

  it("is not editable when readOnly", () => {
    const box = handleBox();
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={{ kind: "readOnly" }}
        ref={box}
        renderImportError={() => null}
      />
    );
    expect(box.current?.view.editable).toBe(false);
    unmount();
  });

  it("offers document zoom in the editing toolbar", () => {
    const select = () =>
      host.querySelector<HTMLSelectElement>(`.${editorClassNames.zoomSelect}`);
    const unmountEdit = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        renderImportError={() => null}
      />
    );
    expect(select()?.value).toBe("fit-width");
    expect(select()?.selectedOptions[0]?.textContent).toBe("Fit");
    unmountEdit();

    const unmountReadOnly = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={{ kind: "readOnly" }}
        renderImportError={() => null}
      />
    );
    expect(select()).toBeNull();
    unmountReadOnly();
  });

  it("keeps zoom internally when it is uncontrolled", () => {
    const changed: unknown[] = [];
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        defaultZoom={0.75}
        onZoomChange={(value) => changed.push(value)}
        renderImportError={() => null}
      />
    );
    const select = host.querySelector<HTMLSelectElement>(
      `.${editorClassNames.zoomSelect}`
    );
    const layer = host.querySelector<HTMLElement>(
      `.${editorClassNames.pageLayer}`
    );
    expect(select?.value).toBe("0.75");
    expect(layer?.style.zoom).toBe("0.75");

    act(() => {
      if (!select) throw new Error("zoom control missing");
      select.value = "1.25";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });

    expect(select?.value).toBe("1.25");
    expect(layer?.style.zoom).toBe("1.25");
    expect(changed).toEqual([1.25]);
    unmount();
  });

  it("publishes the zoom factor to CSS for the panels beside the paper", () => {
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        defaultZoom={0.5}
        renderImportError={() => null}
      />
    );
    const workspace = host.querySelector<HTMLElement>(
      `.${editorClassNames.workspace}`
    );
    const select = host.querySelector<HTMLSelectElement>(
      `.${editorClassNames.zoomSelect}`
    );
    expect(workspace?.style.getPropertyValue("--docx-editor-zoom")).toBe("0.5");

    act(() => {
      if (!select) throw new Error("zoom control missing");
      select.value = "1.5";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });

    expect(workspace?.style.getPropertyValue("--docx-editor-zoom")).toBe("1.5");
    unmount();
  });

  it("reports controlled zoom changes without replacing the prop", () => {
    const changed: unknown[] = [];
    const unmount = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        zoom={1}
        onZoomChange={(value) => changed.push(value)}
        renderImportError={() => null}
      />
    );
    const select = host.querySelector<HTMLSelectElement>(
      `.${editorClassNames.zoomSelect}`
    );
    const layer = host.querySelector<HTMLElement>(
      `.${editorClassNames.pageLayer}`
    );

    act(() => {
      if (!select) throw new Error("zoom control missing");
      select.value = "0.5";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });

    expect(changed).toEqual([0.5]);
    expect(select?.value).toBe("1");
    expect(layer?.style.zoom).toBe("1");
    unmount();
  });

  it("draws the page guides even when readOnly, and not at all when turned off", async () => {
    const guides = () =>
      host.querySelectorAll(`.${editorClassNames.pageGuides}`).length;

    const unmountReadOnly = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={{ kind: "readOnly" }}
        renderImportError={() => null}
      />
    );
    await untilGuides(guides, 1);
    expect(guides()).toBe(1);
    unmountReadOnly();

    const unmountOff = render(
      <DocxEditor
        document={ONE_PARAGRAPH}
        mode={EDITING}
        showPageGuides={false}
        renderImportError={() => null}
      />
    );
    // The frame the guides above were drawn on comes and goes, and draws none of them here
    await nextFrame();
    expect(guides()).toBe(0);
    unmountOff();
  });

  it("renders the text with the fallback fonts handed in by the consumer", () => {
    const runFont = () =>
      host.querySelector<HTMLElement>(`.${editorClassNames.run}`)?.style
        .fontFamily;

    const unmountDefault = render(
      <DocxEditor
        document={BATANG_PARAGRAPH}
        mode={EDITING}
        renderImportError={() => null}
      />
    );
    expect(runFont()).toContain("Noto Serif KR");
    unmountDefault();

    const unmountGiven = render(
      <DocxEditor
        document={BATANG_PARAGRAPH}
        mode={EDITING}
        fontFallbacks={{
          groups: [{ stack: '"Noto Serif Devanagari"', names: ["Batang"] }],
          defaultStack: '"Noto Sans Devanagari"',
          defaultFontName: "Noto Sans Devanagari",
        }}
        renderImportError={() => null}
      />
    );
    // The name the document declared still leads, and only what follows it changed
    expect(runFont()).toBe('"Batang", "Noto Serif Devanagari"');
    unmountGiven();
  });

  // A set written inline is a new object on every render, so reading the prop again would rebuild
  // the whole view under a parent that re-renders for reasons of its own
  it("keeps the fallbacks it mounted with when a new set arrives", () => {
    const run = () =>
      host.querySelector<HTMLElement>(`.${editorClassNames.run}`);
    const mounted: FontFallbacks = {
      groups: [{ stack: '"Noto Serif Devanagari"', names: ["Batang"] }],
      defaultStack: '"Noto Sans Devanagari"',
      defaultFontName: "Noto Sans Devanagari",
    };
    const later: FontFallbacks = {
      groups: [{ stack: '"Noto Serif Armenian"', names: ["Batang"] }],
      defaultStack: '"Noto Sans Armenian"',
      defaultFontName: "Noto Sans Armenian",
    };

    function Host() {
      const [swapped, setSwapped] = useState(false);
      return (
        <>
          <button type="button" onClick={() => setSwapped(true)}>
            Replace
          </button>
          <DocxEditor
            document={BATANG_PARAGRAPH}
            mode={EDITING}
            fontFallbacks={swapped ? later : mounted}
            renderImportError={() => null}
          />
        </>
      );
    }

    const unmount = render(<Host />);
    const drawn = run();
    expect(drawn?.style.fontFamily).toBe('"Batang", "Noto Serif Devanagari"');

    act(() => host.querySelector("button")?.click());
    // The very same element, so the new set rebuilt neither the run nor the view around it
    expect(run()).toBe(drawn);
    expect(run()?.style.fontFamily).toBe('"Batang", "Noto Serif Devanagari"');
    unmount();
  });

  it("marks a run with the language its document records", () => {
    const unmount = render(
      <DocxEditor
        document={readFixture("east-asian.docx")}
        mode={EDITING}
        renderImportError={() => null}
      />
    );
    const runWithText = (text: string) =>
      Array.from(
        host.querySelectorAll<HTMLElement>(`.${editorClassNames.run}`)
      ).find((run) => run.textContent?.startsWith(text));

    expect(runWithText("電車は")?.lang).toBe("ja-JP");
    expect(runWithText("茶还是")?.lang).toBe("zh-CN");
    unmount();
  });

  it("renders the new document when the document changes", () => {
    function Host() {
      const [buffer, setBuffer] = useState(ONE_PARAGRAPH);
      return (
        <>
          <button type="button" onClick={() => setBuffer(OTHER_PARAGRAPH)}>
            Replace
          </button>
          <DocxEditor
            document={buffer}
            mode={EDITING}
            renderImportError={() => null}
          />
        </>
      );
    }

    const unmount = render(<Host />);
    act(() => host.querySelector("button")?.click());
    expect(host.textContent).toContain("other text");
    expect(host.textContent).not.toContain("source");
    unmount();
  });

  describe("the modes", () => {
    const COMMENTING: DocxEditorMode = { kind: "comment", author: AUTHOR };

    /** Selects the whole of the first paragraph's text */
    function selectAll(view: EditorView): void {
      act(() => {
        view.dispatch(
          view.state.tr.setSelection(
            TextSelection.create(view.state.doc, 1, 1 + "source".length)
          )
        );
      });
    }

    function typeInto(view: EditorView): void {
      act(() => {
        view.dispatch(view.state.tr.insertText("typed ", 1));
      });
    }

    it("lets a commenter select text but not change it, with no toolbar to change it from", () => {
      const box = handleBox();
      const unmount = render(
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={COMMENTING}
          ref={box}
          renderImportError={() => null}
        />
      );
      const { view } = attached(box);
      expect(view.editable).toBe(false);
      expect(editingProtection(view.state)).toBe("comments");
      expect(host.querySelector('[role="toolbar"]')).toBeNull();

      selectAll(view);
      expect(view.state.selection.empty).toBe(false);
      expect(view.state.doc.textContent).toBe("source");

      typeInto(view);
      expect(view.state.doc.textContent).toBe("source");
      unmount();
    });

    it("changes mode on the open document without rebuilding the editor", () => {
      const box = handleBox();
      const modes: DocxEditorMode[] = [
        { kind: "readOnly" },
        COMMENTING,
        { kind: "edit", author: AUTHOR },
      ];

      function Host() {
        const [at, setAt] = useState(0);
        return (
          <>
            <button type="button" onClick={() => setAt((index) => index + 1)}>
              Next mode
            </button>
            <DocxEditor
              document={ONE_PARAGRAPH}
              mode={modes[at] ?? { kind: "readOnly" }}
              ref={box}
              renderImportError={() => null}
            />
          </>
        );
      }

      const unmount = render(<Host />);
      const reader = attached(box).view;
      expect(reader.editable).toBe(false);
      typeInto(reader);
      expect(reader.state.doc.textContent).toBe("source");

      act(() => host.querySelector("button")?.click());
      const commenter = attached(box).view;
      expect(commenter).toBe(reader);
      expect(editingProtection(commenter.state)).toBe("comments");
      expect(commenter.editable).toBe(false);

      act(() => host.querySelector("button")?.click());
      const editor = attached(box).view;
      expect(editor).toBe(reader);
      expect(editingProtection(editor.state)).toBe("none");
      expect(editor.editable).toBe(true);
      typeInto(editor);
      expect(editor.state.doc.textContent).toBe("typed source");
      unmount();
    });
  });

  describe("the right click menus", () => {
    /** The rows of whichever menu stands open */
    const menuRows = () =>
      host.querySelectorAll('button[role="menuitem"]').length;

    function mount(mode: DocxEditorMode = EDITING) {
      const box = handleBox();
      const unmount = render(
        <DocxEditor
          document={PARAGRAPH_AND_TABLE}
          mode={mode}
          ref={box}
          renderImportError={() => null}
        />
      );
      const handle = box.current;
      if (!handle) throw new Error("the ref was not attached");
      // jsdom draws nothing, so the spot a right click lands on is answered here: the caret the
      // state already holds, which is what a click on the current selection lands on
      handle.view.posAtCoords = () => ({
        pos: handle.view.state.selection.head,
        inside: -1,
      });
      return unmount;
    }

    /** Right clicks the element, and reports whether the browser's own menu was taken away */
    function rightClick(selector: string): boolean {
      const target = host.querySelector(selector);
      if (!target) throw new Error(`nothing to right click: ${selector}`);
      const event = new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        clientX: 40,
        clientY: 60,
      });
      act(() => {
        target.dispatchEvent(event);
      });
      return event.defaultPrevented;
    }

    it("stands in for the browser's own menu", () => {
      const unmount = mount();
      expect(rightClick("p")).toBe(true);
      expect(menuRows()).toBeGreaterThan(0);
      unmount();
    });

    it("hands the right click back to the browser where the consumer turned them off", () => {
      const unmount = mount({
        kind: "edit",
        author: AUTHOR,
        contextMenus: false,
      });

      // Nothing answered the event, so the browser draws the menu it always would
      expect(rightClick("p")).toBe(false);
      expect(rightClick("td")).toBe(false);
      expect(menuRows()).toBe(0);
      unmount();
    });
  });

  describe("consumer plugins", () => {
    const editedKey = new PluginKey<boolean>("testEdited");
    const EDIT_MARK = "[edited]";

    /**
     * Remembers whether the document has ever been edited, and the first time it is,
     * appends a mark to the end of the last paragraph.
     */
    function editWatcher(): Plugin<boolean> {
      return new Plugin<boolean>({
        key: editedKey,
        state: {
          init: () => false,
          apply: (tr, edited) => edited || tr.docChanged,
        },
        appendTransaction(transactions, _before, after) {
          if (!transactions.some((tr) => tr.docChanged)) return null;
          if (after.doc.textContent.includes(EDIT_MARK)) return null;
          return after.tr.insertText(EDIT_MARK, after.doc.content.size - 1);
        },
      });
    }

    /** Writes a mark instead of splitting the paragraph, which is what Enter does by default */
    const writeReturnMark: Command = (state, dispatch) => {
      if (dispatch) dispatch(state.tr.insertText("↵"));
      return true;
    };

    /** Presses Enter on the editing surface the way a browser would */
    function pressEnter(view: EditorView) {
      act(() => {
        view.dom.dispatchEvent(
          new KeyboardEvent("keydown", {
            key: "Enter",
            bubbles: true,
            cancelable: true,
          })
        );
      });
    }

    /** The editor that the ref handed back. Fails the test when it was never attached */
    it("lets a plugin keep its own state and append transactions", () => {
      const box = handleBox();
      const unmount = render(
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={EDITING}
          plugins={[editWatcher()]}
          ref={box}
          renderImportError={() => null}
        />
      );

      const editor = attached(box);
      expect(editedKey.getState(editor.view.state)).toBe(false);

      act(() => {
        editor.view.dispatch(editor.view.state.tr.insertText("edited ", 1));
      });

      expect(editedKey.getState(editor.view.state)).toBe(true);
      expect(host.textContent).toContain(`edited source${EDIT_MARK}`);
      unmount();
    });

    it("gives a consumer keymap the first look, ahead of the built-in one", () => {
      const builtIn = handleBox();
      const unmountBuiltIn = render(
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={EDITING}
          ref={builtIn}
          renderImportError={() => null}
        />
      );
      const withBuiltInEnter = attached(builtIn).view;
      pressEnter(withBuiltInEnter);
      expect(withBuiltInEnter.state.doc.childCount).toBe(2);
      unmountBuiltIn();

      const overridden = handleBox();
      const unmountOverridden = render(
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={EDITING}
          plugins={[keymap({ Enter: writeReturnMark })]}
          ref={overridden}
          renderImportError={() => null}
        />
      );
      const withConsumerEnter = attached(overridden).view;
      pressEnter(withConsumerEnter);
      expect(withConsumerEnter.state.doc.childCount).toBe(1);
      expect(host.textContent).toContain("↵source");
      unmountOverridden();
    });

    // A consumer plugin is not a way past the mode: what it dispatches is judged like anything else
    it("hands the plugins to a readOnly editor as well, whose refusals they share", () => {
      const box = handleBox();
      const unmount = render(
        <DocxEditor
          document={ONE_PARAGRAPH}
          mode={{ kind: "readOnly" }}
          plugins={[editWatcher()]}
          ref={box}
          renderImportError={() => null}
        />
      );

      const editor = attached(box);
      expect(editor.view.editable).toBe(false);
      expect(editedKey.getState(editor.view.state)).toBe(false);

      act(() => {
        editor.view.dispatch(editor.view.state.tr.insertText("edited ", 1));
      });
      expect(editedKey.getState(editor.view.state)).toBe(false);
      expect(editor.view.state.doc.textContent).toBe("source");
      unmount();
    });

    // The state built for the new document is the one place where reading the prop again would
    // show, so this empties the array and swaps the document in the same render
    it("keeps the plugins it mounted with when the array changes later", () => {
      const box = handleBox();

      function Host() {
        const [swapped, setSwapped] = useState(false);
        return (
          <>
            <button type="button" onClick={() => setSwapped(true)}>
              Replace
            </button>
            <DocxEditor
              document={swapped ? OTHER_PARAGRAPH : ONE_PARAGRAPH}
              mode={EDITING}
              plugins={swapped ? [] : [editWatcher()]}
              ref={box}
              renderImportError={() => null}
            />
          </>
        );
      }

      const unmount = render(<Host />);
      act(() => host.querySelector("button")?.click());

      const editor = attached(box);
      expect(editedKey.getState(editor.view.state)).toBe(false);
      act(() => {
        editor.view.dispatch(editor.view.state.tr.insertText("edited ", 1));
      });
      expect(host.textContent).toContain(`edited other text${EDIT_MARK}`);
      unmount();
    });
  });
});
