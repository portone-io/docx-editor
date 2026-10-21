// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeNotesDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import {
  createEditorView,
  editorStateForSession,
} from "../editor/createEditor";
import { documentOf, storyDocument } from "../editor/editorDocument";
import { noteExtensions, noteHost } from "../editor/notes/noteSurface";
import { createStoryView, type StoryView } from "../editor/stories/storyView";
import { storyKey } from "../schema/stories";
import { editorClassNames } from "./classNames";
import { DEFAULT_FONT_FALLBACKS } from "./fontStack";

const FOOTNOTE = storyKey("footnote", "2");

const OPENS_ON_REFERENCE =
  '<w:p><w:r><w:footnoteReference w:id="2"/></w:r>' +
  '<w:r><w:t xml:space="preserve">Text</w:t></w:r>' +
  '<w:r><w:endnoteReference w:id="3"/></w:r></w:p>';

/**
 * prosemirror-view puts a separator image in front of an inline atom, which Chrome composes
 * around wrongly and which opens an empty line in front of a block-level page space. The same
 * class also marks the image that keeps a line ending in an atom tall enough for the caret, and
 * the cursor wrapper a composition starts in, and both of those have to stay on screen.
 */
describe("the separator rule in editor.css", () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "editor.css"),
    "utf8"
  );
  const selector =
    /(?:^|[;}])([^;{}]*img\.ProseMirror-separator[^{}]*)\{[^}]*display:\s*none/
      .exec(css.replace(/\/\*[\s\S]*?\*\//g, ""))?.[1]
      ?.replace(/\s+/g, " ")
      .trim() ?? "";

  const opened: { main: EditorView | null; story: StoryView | null } = {
    main: null,
    story: null,
  };

  afterEach(() => {
    opened.story?.destroy();
    opened.main?.destroy();
    opened.story = null;
    opened.main = null;
    document.body.replaceChildren();
  });

  function mount(): HTMLElement {
    return document.body.appendChild(document.createElement("div"));
  }

  function body(xml?: string): EditorView {
    opened.main = createEditorView({
      mount: mount(),
      state: editorStateForSession(importDocx(makeNotesDocx(xml))),
      onStateChange: () => {},
    });
    return opened.main;
  }

  /** The footnote opened for editing, the way a row does it (`ui/notes/StoryRow`) */
  function footnote(): EditorView {
    const main = body();
    opened.story = createStoryView({
      mount: mount(),
      host: noteHost(main, () => {}),
      key: FOOTNOTE,
      document: storyDocument(
        documentOf(main.state),
        documentOf(main.state).geometry
      ),
      fontFallbacks: DEFAULT_FONT_FALLBACKS,
      extensions: noteExtensions(main, FOOTNOTE, () => "1"),
    });
    return opened.story.view;
  }

  function separator(found: Element | null | undefined): Element {
    if (!found?.matches("img.ProseMirror-separator")) {
      throw new Error("no separator there");
    }
    return found;
  }

  it("hides the separator in front of the number a note opens with", () => {
    const view = footnote();
    const leading = separator(
      view.dom.querySelector(`.${editorClassNames.noteMark}`)
        ?.previousElementSibling
    );

    expect(leading.matches(selector)).toBe(true);
  });

  it("hides the separator in front of a reference a body line opens with", () => {
    const view = body(OPENS_ON_REFERENCE);
    const leading = separator(
      view.dom.querySelector(`.${editorClassNames.noteReference}`)
        ?.previousElementSibling
    );

    expect(leading.matches(selector)).toBe(true);
  });

  it("keeps the separator that holds a line ending in an atom open", () => {
    const view = body(OPENS_ON_REFERENCE);
    const references = view.dom.querySelectorAll(
      `.${editorClassNames.noteReference}`
    );
    const trailing = separator(
      references[references.length - 1]?.nextElementSibling
    );

    expect(trailing.matches(selector)).toBe(false);
  });

  it("keeps the cursor wrapper a composition starts in", () => {
    const view = body();
    const beforeReference = 1 + "Text".length;
    view.dispatch(
      view.state.tr
        .setSelection(TextSelection.create(view.state.doc, beforeReference))
        .setStoredMarks([])
    );
    view.focus();
    view.dom.dispatchEvent(
      new CompositionEvent("compositionstart", { bubbles: true })
    );
    const wrapper = separator(view.dom.querySelector("img[mark-placeholder]"));

    expect(wrapper.nextElementSibling).not.toBeNull();
    expect(wrapper.matches(selector)).toBe(false);
  });
});
