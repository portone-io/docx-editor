// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import {
  LETTER_SECT_PR,
  makeNotesDocx,
  NOTE_BODY,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { storyOf, storyText } from "../../docx/story";
import { storyKey } from "../../schema/stories";
import { DEFAULT_FONT_FALLBACKS } from "../../styles/fontStack";
import { readContextOf } from "../clipboard/readContext";
import { undo } from "../commands/historyCommands";
import { createEditorView, editorStateForSession } from "../createEditor";
import { documentOf, storyDocument } from "../editorDocument";
import { footnoteExtensions, footnoteHost } from "../notes/footnoteSurface";
import { createStoryView, type StoryHost, type StoryView } from "./storyView";

const FOOTNOTE = storyKey("footnote", "2");

const opened: { main: EditorView | null; story: StoryView | null } = {
  main: null,
  story: null,
};

afterEach(() => {
  opened.story?.destroy();
  opened.main?.destroy();
  opened.story = null;
  opened.main = null;
});

function mainView(body: string = NOTE_BODY): EditorView {
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(importDocx(makeNotesDocx(body))),
    onStateChange: () => {},
  });
  opened.main = view;
  return view;
}

/** The footnote opened for editing, the way a row does it (`ui/notes/StoryRow`) */
function openFootnote(
  main: EditorView,
  host: StoryHost = footnoteHost(main, () => {})
): StoryView {
  const story = createStoryView({
    mount: document.createElement("div"),
    host,
    key: FOOTNOTE,
    document: storyDocument(
      documentOf(main.state),
      documentOf(main.state).geometry
    ),
    fontFallbacks: DEFAULT_FONT_FALLBACKS,
    extensions: footnoteExtensions(main, "2", () => "1"),
  });
  opened.story = story;
  return story;
}

/** The key `Mod` stands for here, read the way the keymap reads it */
const MOD = /Mac/.test(navigator.platform) ? "metaKey" : "ctrlKey";

function pressed(view: EditorView, key: string, mod = true): boolean {
  const event = new KeyboardEvent("keydown", {
    key,
    [MOD]: mod,
    bubbles: true,
    cancelable: true,
  });
  return view.someProp("handleKeyDown", (run) => run(view, event)) === true;
}

function footnoteIn(doc: PMNode): PMNode | null {
  return storyOf(doc, FOOTNOTE);
}

describe("the view one story is edited in", () => {
  it("writes a keystroke into the host and nowhere else", () => {
    const main = mainView();
    const bodyBefore = main.state.doc.textContent;
    const story = openFootnote(main);

    story.view.dispatch(story.view.state.tr.insertText("!", 2));

    expect(storyText(footnoteIn(main.state.doc))).toBe(
      "!Footnote body\nSecond line"
    );
    expect(main.state.doc.textContent).toBe(bodyBefore);
    // The document holds the one history both surfaces share, so the edit is there to take back
    expect(undo(main.state)).toBe(true);
  });

  it("takes an edit back from inside the story through the host's history", () => {
    const main = mainView();
    const story = openFootnote(main);
    story.view.dispatch(story.view.state.tr.insertText("!", 2));
    expect(storyText(footnoteIn(main.state.doc))).toContain("!");

    expect(pressed(story.view, "z")).toBe(true);
    story.sync();

    expect(storyText(footnoteIn(main.state.doc))).toBe(
      "Footnote body\nSecond line"
    );
    expect(storyText(story.view.state.doc)).toBe("Footnote body\nSecond line");
  });

  it("keeps an open composition when the main story says the same thing", () => {
    const main = mainView();
    const story = openFootnote(main);
    story.view.dispatch(story.view.state.tr.insertText("!", 2));
    const held = story.view.state;

    story.sync();

    // Nothing was replaced, so the very state the keystroke left is the one still standing and
    // whatever the browser is composing in it was never redrawn
    expect(story.view.state).toBe(held);
  });

  it("replaces the content when an undo changes the story", () => {
    const main = mainView();
    const story = openFootnote(main);
    story.view.dispatch(story.view.state.tr.insertText("!", 2));
    const typed = story.view.state.doc;

    main.dispatch(main.state.tr.setDocAttribute("stories", {}));
    main.dispatch(
      main.state.tr.setDocAttribute("stories", {
        [FOOTNOTE]: typed.type.schema.nodes.doc
          .create(null, [
            typed.type.schema.nodes.paragraph.create(
              null,
              typed.type.schema.text("Rewritten")
            ),
          ])
          .toJSON(),
      })
    );
    story.sync();

    expect(storyText(story.view.state.doc)).toBe("Rewritten");
    expect(story.view.state.doc).not.toBe(typed);
  });

  it("keeps its state when the host refuses a write", () => {
    const main = mainView();
    const refusing: StoryHost = {
      ...footnoteHost(main, () => {}),
      write: () => false,
    };
    const story = openFootnote(main, refusing);
    const before = story.view.state;

    story.view.dispatch(story.view.state.tr.insertText("!", 2));

    expect(story.view.state.doc).toBe(before.doc);
    expect(storyText(footnoteIn(main.state.doc))).toBe(
      "Footnote body\nSecond line"
    );
  });

  it("keeps the state an edit left when the host has nothing to write", () => {
    const main = mainView();
    const story = openFootnote(main);
    const before = story.view.state;

    // The same words over themselves: the story ends up saying what it already said, so the
    // document writes nothing, which is not the same as turning the edit down
    story.view.dispatch(story.view.state.tr.insertText("Footnote body", 2, 15));

    // The view stands on what the edit left rather than being wound back onto what it was,
    // which is what breaks an open composition
    expect(story.view.state).not.toBe(before);
    expect(story.view.state.doc).not.toBe(before.doc);
    expect(storyText(story.view.state.doc)).toBe("Footnote body\nSecond line");
  });

  it("reads a paste against the paper its snapshot names", () => {
    const main = mainView(NOTE_BODY + LETTER_SECT_PR);
    const story = openFootnote(main);

    const context = readContextOf(story.view.state, document);

    expect(context.geometry).toBe(documentOf(story.view.state).geometry);
    expect(context.geometry).toEqual(documentOf(main.state).geometry);
    // A story has nowhere of its own to write a list definition
    expect(context.numbering.canCreate).toBe(false);
  });

  it("registers the view that holds the caret and lets go on the way out", () => {
    const main = mainView();
    const registered: (EditorView | null)[] = [];
    const story = openFootnote(
      main,
      footnoteHost(main, (_key, view) => registered.push(view))
    );

    story.view.dom.dispatchEvent(new FocusEvent("focus"));
    expect(registered).toEqual([story.view]);

    story.destroy();
    opened.story = null;
    expect(registered).toEqual([story.view, null]);
  });
});

describe("the state one story is edited in", () => {
  it("keeps no history of its own, since the document holds the one both share", () => {
    const main = mainView();
    const story = openFootnote(main);
    story.view.dispatch(story.view.state.tr.insertText("!", 2));

    // Taking that keystroke back is the document's to do, so the story's own state has nothing
    expect(undo(story.view.state)).toBe(false);
    expect(undo(main.state)).toBe(true);
  });
});
