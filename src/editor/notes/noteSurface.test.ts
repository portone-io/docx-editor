// @vitest-environment jsdom
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import {
  makeNotesDocx,
  NOTE_BODY,
  TINY_PNG_DATA_URL,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { storyOf, storyText } from "../../docx/story";
import { docxSchema } from "../../schema";
import type { EditingProtection } from "../../schema/protection";
import { storyKey } from "../../schema/stories";
import { DEFAULT_FONT_FALLBACKS } from "../../styles/fontStack";
import { normalizePasted } from "../clipboard/normalizers";
import { insertFootnote } from "../commands/footnoteCommands";
import { listRefOf, toggleNumberedList } from "../commands/listCommands";
import { createEditorView, editorStateForSession } from "../createEditor";
import { documentOf, storyDocument } from "../editorDocument";
import { requestedNote } from "../plugins/noteNavigation";
import { createStoryView, type StoryView } from "../stories/storyView";
import { footnoteExtensions, footnoteHost } from "./footnoteSurface";

const FOOTNOTE = storyKey("footnote", "2");

/** The footnote reference of `NOTE_BODY`, standing inside a control that takes no edit */
const LOCKED_REFERENCE =
  "<w:p><w:sdt>" +
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
  "<w:sdtContent>" +
  '<w:r><w:t xml:space="preserve">Text</w:t></w:r>' +
  '<w:r><w:footnoteReference w:id="2"/></w:r>' +
  "</w:sdtContent></w:sdt></w:p>";

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

function mainView(
  body: string = NOTE_BODY,
  protection: EditingProtection = "none"
): EditorView {
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(importDocx(makeNotesDocx(body)), {
      protection,
      author: { id: "me", name: "Me" },
    }),
    onStateChange: () => {},
  });
  opened.main = view;
  return view;
}

function openNote(main: EditorView, id = "2", label = "1"): StoryView {
  const story = createStoryView({
    mount: document.createElement("div"),
    host: footnoteHost(main, () => {}),
    key: storyKey("footnote", id),
    document: storyDocument(
      documentOf(main.state),
      documentOf(main.state).geometry
    ),
    fontFallbacks: DEFAULT_FONT_FALLBACKS,
    extensions: footnoteExtensions(main, id, () => label),
  });
  opened.story = story;
  return story;
}

function pressBackspace(view: EditorView): boolean {
  const event = new KeyboardEvent("keydown", {
    key: "Backspace",
    bubbles: true,
    cancelable: true,
  });
  return view.someProp("handleKeyDown", (run) => run(view, event)) === true;
}

/** Mod is the platform's own modifier, read the way `prosemirror-keymap` reads it */
const ON_MAC =
  typeof navigator !== "undefined" && /Mac/.test(navigator.platform);

function pressInsertFootnote(view: EditorView): boolean {
  const event = new KeyboardEvent("keydown", {
    key: "f",
    altKey: true,
    ctrlKey: !ON_MAC,
    metaKey: ON_MAC,
    bubbles: true,
    cancelable: true,
  });
  return view.someProp("handleKeyDown", (run) => run(view, event)) === true;
}

function paragraph(...content: readonly PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(
    null,
    Fragment.fromArray([...content])
  );
}

describe("what a note takes", () => {
  it("refuses typing in a footnote under comment mode and leaves the story as it was", () => {
    const main = mainView(NOTE_BODY, "comments");
    const story = openNote(main);
    const before = story.view.state.doc;

    story.view.dispatch(story.view.state.tr.insertText("!", 2));

    expect(story.view.state.doc).toBe(before);
    expect(storyText(storyOf(main.state.doc, FOOTNOTE))).toBe(
      "Footnote body\nSecond line"
    );
    expect(story.view.editable).toBe(false);
  });

  it("refuses typing in a footnote whose reference stands in locked content", () => {
    const main = mainView(LOCKED_REFERENCE);
    const story = openNote(main);
    const before = story.view.state.doc;

    story.view.dispatch(story.view.state.tr.insertText("!", 2));

    expect(story.view.state.doc).toBe(before);
    expect(storyText(storyOf(main.state.doc, FOOTNOTE))).toBe(
      "Footnote body\nSecond line"
    );
  });

  it("binds no footnote insertion inside a footnote, though the body has the key", () => {
    const main = mainView();
    const story = openNote(main);
    const before = main.state.doc;

    // A note may not hold a note (5.3 of the notes plan), and the key is bound only where the
    // surface takes what it puts in (`editor/stories/storyState`), so a note is bound none
    expect(pressInsertFootnote(story.view)).toBe(false);
    expect(main.state.doc).toBe(before);

    main.dispatch(
      main.state.tr.setSelection(TextSelection.create(main.state.doc, 1))
    );
    expect(pressInsertFootnote(main)).toBe(true);
    expect(main.state.doc).not.toBe(before);
  });

  it("drops an image or a new link pasted into a footnote", () => {
    const main = mainView();
    const story = openNote(main);
    const linked = docxSchema
      .text("linked")
      .mark([docxSchema.marks.link.create({ xml: "<w:hyperlink>" })]);
    const image = docxSchema.nodes.image.create({
      src: TINY_PNG_DATA_URL,
      extent: { cx: 952500, cy: 952500 },
    });

    const normalized = normalizePasted(
      {
        slice: new Slice(
          Fragment.from(paragraph(linked, image, docxSchema.text(" plain"))),
          0,
          0
        ),
        newLists: new Map(),
      },
      story.view.state,
      false,
      footnoteExtensions(main, "2", () => "1").normalizers
    );

    const first = normalized.slice.content.firstChild;
    expect(first?.textContent).toBe("linked plain");
    expect(
      first?.children.some((child) => child.type === docxSchema.nodes.image)
    ).toBe(false);
    expect(
      first?.children.some((child) =>
        child.marks.some((mark) => mark.type === docxSchema.marks.link)
      )
    ).toBe(false);
  });

  it("drops a footnote reference pasted into a footnote", () => {
    const main = mainView();
    const story = openNote(main);
    const reference = docxSchema.nodes.noteReference.create({
      kind: "footnote",
      id: "2",
      label: "1",
      referenceXml: '<w:footnoteReference w:id="2"/>',
    });
    const called = storyOf(main.state.doc, FOOTNOTE);
    if (called === null) throw new Error("the document holds no footnote 2");

    const normalized = normalizePasted(
      {
        slice: new Slice(
          Fragment.from(paragraph(docxSchema.text("called"), reference)),
          0,
          0
        ),
        newLists: new Map(),
        noteStories: new Map([[FOOTNOTE, called]]),
      },
      story.view.state,
      false,
      footnoteExtensions(main, "2", () => "1").normalizers
    );

    const first = normalized.slice.content.firstChild;
    expect(first?.textContent).toBe("called");
    expect(
      first?.children.some(
        (child) => child.type === docxSchema.nodes.noteReference
      )
    ).toBe(false);
    expect(normalized.newStories?.size ?? 0).toBe(0);
  });

  it("does not start a new list inside a footnote", () => {
    const main = mainView();
    const story = openNote(main);

    expect(toggleNumberedList(story.view.state)).toBe(false);

    const listed = docxSchema.nodes.paragraph.create(
      {
        pPr:
          "<w:pPr><w:numPr>" +
          '<w:ilvl w:val="0"/><w:numId w:val="99"/>' +
          "</w:numPr></w:pPr>",
        format: { numbering: { numId: 99, ilvl: 0 } },
      },
      docxSchema.text("pasted")
    );
    const normalized = normalizePasted(
      { slice: new Slice(Fragment.from(listed), 0, 0), newLists: new Map() },
      story.view.state,
      false,
      footnoteExtensions(main, "2", () => "1").normalizers
    );

    const first = normalized.slice.content.firstChild;
    expect(first === null ? null : listRefOf(first)).toBeNull();
    expect(normalized.newLists.size).toBe(0);
  });

  it("deletes an empty footnote and its reference on Backspace at its start", () => {
    const main = mainView();
    const at = main.state.doc.content.size - 1;
    main.dispatch(
      main.state.tr.setSelection(TextSelection.create(main.state.doc, at))
    );
    insertFootnote(main.state, (tr) => main.dispatch(tr));
    const asked = requestedNote(main.state);
    if (asked === null) throw new Error("no footnote was asked to open");
    const id = asked.key.slice(storyKey("footnote", "").length);
    const references = () => {
      let found = 0;
      main.state.doc.descendants((node) => {
        if (node.type === docxSchema.nodes.noteReference) found += 1;
        return true;
      });
      return found;
    };
    expect(references()).toBe(3);
    expect(storyOf(main.state.doc, asked.key)).not.toBeNull();

    const story = openNote(main, id, "2");
    story.view.dispatch(
      story.view.state.tr.setSelection(
        TextSelection.create(story.view.state.doc, 1)
      )
    );
    expect(pressBackspace(story.view)).toBe(true);

    expect(references()).toBe(2);
    expect(storyOf(main.state.doc, asked.key)).toBeNull();
    // The caret is handed back to where the reference stood
    expect(main.state.selection.from).toBe(at);
  });

  it("keeps a footnote holding an image alone on Backspace at its start", () => {
    const main = mainView();
    const story = openNote(main);
    const image = docxSchema.nodes.image.create({
      src: TINY_PNG_DATA_URL,
      extent: { cx: 952500, cy: 952500 },
    });
    // The note is left holding its number and the image, which spells no text at all
    story.view.dispatch(
      story.view.state.tr.delete(2, story.view.state.doc.content.size - 1)
    );
    story.view.dispatch(story.view.state.tr.insert(2, image));
    expect(storyText(storyOf(main.state.doc, FOOTNOTE))).toBe("");
    story.view.dispatch(
      story.view.state.tr.setSelection(
        TextSelection.create(story.view.state.doc, 1)
      )
    );

    expect(pressBackspace(story.view)).toBe(false);

    expect(storyOf(main.state.doc, FOOTNOTE)).not.toBeNull();
  });

  it("does nothing on Backspace at the start of a footnote that holds text", () => {
    const main = mainView();
    const story = openNote(main);
    story.view.dispatch(
      story.view.state.tr.setSelection(
        TextSelection.create(story.view.state.doc, 1)
      )
    );

    expect(pressBackspace(story.view)).toBe(false);
    expect(storyText(storyOf(main.state.doc, FOOTNOTE))).toBe(
      "Footnote body\nSecond line"
    );
  });
});
