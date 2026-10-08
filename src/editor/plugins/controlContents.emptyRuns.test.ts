// @vitest-environment jsdom
/**
 * Writing into a run holding no characters, the blank a form leaves to be filled in: what is
 * written takes the run's formatting and the run goes, and inside an open content control the
 * text goes into the control.
 *
 * The run under test is the one a server writing such a form leaves in a table cell and inside a
 * control: a yellow highlight over an empty `w:t`.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { Fragment, Slice } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { ReplaceAroundStep } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { commentedDocx } from "../../__testing__/comments";
import { documentXmlOf, makeDocx } from "../../__testing__/docx";
import {
  openComposition,
  posOfText,
  rangeOfText,
  select,
} from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { docxSchema } from "../../schema";
import { wrappersOf } from "../../schema/wrappers";
import { toggleBold } from "../commands/formatting/editing";
import { activeTextBackground } from "../commands/formatting/reading";
import { undo } from "../commands/historyCommands";
import {
  createEditorState,
  createEditorView,
  editorStateForSession,
} from "../createEditor";
import { docxKeymap } from "./keymap";

const HIGHLIGHT = '<w:rPr><w:highlight w:val="yellow"/></w:rPr>';

const SLOT = `<w:r>${HIGHLIGHT}<w:t xml:space="preserve"></w:t></w:r>`;

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** The run typed text goes out as once it has filled the blank */
const filledRun = (text: string) =>
  `<w:r>${HIGHLIGHT}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const CELL =
  '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:p>${run("Name")}</w:p></w:tc><w:tc><w:p>${SLOT}</w:p></w:tc></w:tr>` +
  "</w:tbl><w:p/>";

function control(
  content: string,
  { id = 5, tag = "SIGNER", props = "" } = {}
): string {
  return (
    `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="${id}"/>${props}</w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

const lock = (val: string) => `<w:lock w:val="${val}"/>`;

/** A label, then the control holding nothing but the blank */
const IN_CONTROL = (props = "") =>
  `<w:p>${run("Signer: ")}${control(SLOT, { props })}</w:p>`;

interface Opened {
  state: EditorState;
  session: ReturnType<typeof importDocx>["session"];
}

function open(body: string): Opened {
  const opened = importDocx(makeDocx(body));
  return { state: editorStateForSession(opened), session: opened.session };
}

function emptyRuns(doc: PMNode): number[] {
  const found: number[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "emptyRun") found.push(pos);
    return true;
  });
  return found;
}

function firstEmptyRun(doc: PMNode): number {
  const [pos] = emptyRuns(doc);
  if (pos === undefined) throw new Error("no empty run");
  return pos;
}

/** The caret just after the first blank, where a press on it puts the caret */
function afterSlot(state: EditorState): EditorState {
  const pos = firstEmptyRun(state.doc) + 1;
  return select(state, pos, pos);
}

/** Types each character at the selection, one transaction each, the way a keyboard does */
function type(state: EditorState, text: string): EditorState {
  let next = state;
  for (const char of text) next = next.apply(next.tr.insertText(char));
  return next;
}

function controlsOf(node: PMNode): readonly Mark[] {
  return wrappersOf(node, docxSchema.marks.sdt);
}

/** The text the document's inline controls hold, one entry per control */
function controlTexts(doc: PMNode): string[] {
  const texts = new Map<Mark, string>();
  doc.descendants((node) => {
    if (!node.isText) return true;
    for (const mark of controlsOf(node)) {
      const known = [...texts.keys()].find((other) => other.eq(mark));
      const key = known ?? mark;
      texts.set(key, (texts.get(key) ?? "") + (node.text ?? ""));
    }
    return true;
  });
  return [...texts.values()];
}

function highlightOf(node: PMNode | null | undefined): unknown {
  const run = node?.marks.find((mark) => mark.type === docxSchema.marks.run);
  const format: unknown = run?.attrs.format;
  return typeof format === "object" && format !== null && "highlight" in format
    ? format.highlight
    : undefined;
}

function textNode(doc: PMNode, text: string): PMNode | null {
  return doc.nodeAt(posOfText(doc, text) - 1);
}

describe("typing into a blank in a table cell", () => {
  it.each([
    ["after it", 1],
    ["before it", 0],
  ])("with the caret %s fills it with highlighted text", (_, offset) => {
    const opened = open(CELL);
    const pos = firstEmptyRun(opened.state.doc) + offset;
    const typed = type(select(opened.state, pos, pos), "Kim");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(highlightOf(textNode(typed.doc, "Kim"))).toBe("yellow");
    const xml = documentXmlOf(typed.doc, opened.session);
    expect(xml).toContain(`<w:tc><w:p>${filledRun("Kim")}</w:p></w:tc>`);
  });

  it("takes the blank back with the text in one undo", () => {
    const opened = open(CELL);
    const typed = type(afterSlot(opened.state), "K");
    let undone: EditorState | null = null;
    undo(typed, (tr) => {
      undone = typed.apply(tr);
    });
    expect((undone as EditorState | null)?.doc.eq(opened.state.doc)).toBe(true);
  });

  it("is filled by text typed after a formatting change at the caret", () => {
    const opened = open(CELL);
    let state = afterSlot(opened.state);
    toggleBold(state, (tr) => {
      state = state.apply(tr);
    });
    const typed = type(state, "K");

    expect(emptyRuns(typed.doc)).toEqual([]);
    const node = textNode(typed.doc, "K");
    expect(highlightOf(node)).toBe("yellow");
    expect(documentXmlOf(typed.doc, opened.session)).toContain("<w:b/>");
  });

  it("is filled by a paste, which takes its formatting", () => {
    const opened = open(CELL);
    const state = afterSlot(opened.state);
    const marks = state.doc.resolve(state.selection.from).marks();
    const pasted = state.apply(
      state.tr.replaceSelection(
        new Slice(Fragment.from(docxSchema.text("Kim", marks)), 0, 0)
      )
    );

    expect(emptyRuns(pasted.doc)).toEqual([]);
    expect(highlightOf(textNode(pasted.doc, "Kim"))).toBe("yellow");
  });

  it("is not filled by Enter, which writes no text", () => {
    const opened = open(CELL);
    let state = afterSlot(opened.state);
    docxKeymap.Enter(state, (tr) => {
      state = state.apply(tr);
    });
    expect(emptyRuns(state.doc)).toHaveLength(1);
  });

  it("reads its highlight at the caret, as the toolbar does", () => {
    const opened = open(CELL);
    expect(activeTextBackground(afterSlot(opened.state))).toBe("#ffff00");
  });
});

describe("a blank beside text", () => {
  it("is filled by text typed after it, ahead of the text that follows", () => {
    const opened = open(`<w:p>${SLOT}${run("tail")}</w:p>`);
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(highlightOf(textNode(typed.doc, "K"))).toBe("yellow");
  });

  it("is left alone by a caret after the label, which carries on the label", () => {
    const opened = open(`<w:p>${run("Label: ")}${SLOT}</w:p>`);
    const pos = firstEmptyRun(opened.state.doc);
    const typed = type(select(opened.state, pos, pos), "K");

    expect(emptyRuns(typed.doc)).toHaveLength(1);
    expect(typed.doc.textContent).toBe("Label: K");
    expect(highlightOf(textNode(typed.doc, "Label: K"))).toBeUndefined();
  });

  it("is filled by a paste of several paragraphs", () => {
    const opened = open(`<w:p>${SLOT}${run("tail")}</w:p>`);
    const state = afterSlot(opened.state);
    const paragraph = (text: string) =>
      docxSchema.nodes.paragraph.create(null, docxSchema.text(text));
    const pasted = state.apply(
      state.tr.replaceSelection(
        new Slice(Fragment.from([paragraph("one"), paragraph("two")]), 1, 1)
      )
    );

    expect(pasted.doc.textContent).toBe("onetwotail");
    expect(emptyRuns(pasted.doc)).toEqual([]);
  });

  it("is filled by a step that puts text in front of a gap it carries over", () => {
    // The shape a paste ending deeper than the paragraph it lands in is fitted as
    const opened = open(`<w:p>${SLOT}${run("tail")}</w:p>`);
    const state = afterSlot(opened.state);
    const at = state.selection.from;
    const end = at + "tail".length;
    const marks = state.doc.resolve(at).marks();
    const step = new ReplaceAroundStep(
      at,
      end,
      at,
      end,
      new Slice(Fragment.from(docxSchema.text("one", marks)), 0, 0),
      3
    );
    const pasted = state.apply(state.tr.step(step));

    expect(pasted.doc.textContent).toBe("onetail");
    expect(emptyRuns(pasted.doc)).toEqual([]);
  });
});

describe("a blank that is all an inline control holds", () => {
  it("takes what is typed after it into the control, keystroke after keystroke", () => {
    const opened = open(IN_CONTROL());
    const typed = type(afterSlot(opened.state), "Kim");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(controlTexts(typed.doc)).toEqual(["Kim"]);
    expect(highlightOf(textNode(typed.doc, "Kim"))).toBe("yellow");
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      `<w:sdtContent>${filledRun("Kim")}</w:sdtContent>`
    );
  });

  it("takes text into a control locked against deletion alone, which stays", () => {
    const opened = open(IN_CONTROL(lock("sdtLocked")));
    const typed = type(afterSlot(opened.state), "Kim");

    expect(controlTexts(typed.doc)).toEqual(["Kim"]);
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      `${lock("sdtLocked")}</w:sdtPr><w:sdtContent>${filledRun("Kim")}`
    );
  });

  it("keeps its blank where the lock shuts the contents, and the text stands beside it", () => {
    const opened = open(IN_CONTROL(lock("sdtContentLocked")));
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toHaveLength(1);
    expect(controlTexts(typed.doc)).toEqual([]);
    expect(typed.doc.textContent).toBe("Signer: K");
  });

  it("gives the text beside it no formatting where the lock keeps it from being filled", () => {
    const opened = open(IN_CONTROL(lock("sdtContentLocked")));
    const typed = type(afterSlot(opened.state), "K");

    expect(highlightOf(textNode(typed.doc, "K"))).toBeUndefined();
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      `</w:sdt>${run("K")}`
    );
  });

  it("stops claiming to show its placeholder once it is filled", () => {
    const opened = open(IN_CONTROL("<w:showingPlcHdr/>"));
    const typed = type(afterSlot(opened.state), "K");

    expect(controlTexts(typed.doc)).toEqual(["K"]);
    expect(documentXmlOf(typed.doc, opened.session)).not.toContain(
      "showingPlcHdr"
    );
  });

  it("goes with the blank where the control says it goes with any edit", () => {
    const opened = open(IN_CONTROL("<w:temporary/>"));
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(controlTexts(typed.doc)).toEqual([]);
    const xml = documentXmlOf(typed.doc, opened.session);
    expect(xml).not.toContain("<w:sdt>");
    expect(xml).toContain(filledRun("K"));
  });
});

describe("a temporary control locked against deletion holding a blank", () => {
  it("keeps the control, which its lock says may not go, around the text", () => {
    const opened = open(IN_CONTROL(`<w:temporary/>${lock("sdtLocked")}`));
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(controlTexts(typed.doc)).toEqual(["K"]);
  });
});

describe("a blank beside a blank of other formatting", () => {
  const GREEN_SLOT = SLOT.replace("yellow", "green");

  it("stays where the blank beside it is filled", () => {
    const opened = open(`<w:p>${GREEN_SLOT}${SLOT}</w:p>`);
    const pos = (emptyRuns(opened.state.doc)[1] ?? 0) + 1;
    const typed = type(select(opened.state, pos, pos), "K");

    expect(emptyRuns(typed.doc)).toHaveLength(1);
    expect(highlightOf(textNode(typed.doc, "K"))).toBe("yellow");
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      '<w:highlight w:val="green"/>'
    );
  });
});

describe("a blank beside a control", () => {
  it("stays where the control's placeholder is typed over", () => {
    const placeholder = "[    ]";
    const opened = open(`<w:p>${control(run(placeholder))}${SLOT}</w:p>`);
    const { from, to } = rangeOfText(opened.state.doc, placeholder);
    const typed = type(select(opened.state, from, to), "12");

    expect(controlTexts(typed.doc)).toEqual(["12"]);
    expect(emptyRuns(typed.doc)).toHaveLength(1);
  });

  it("is set aside by a selection running past the control to the end of the line", () => {
    const placeholder = "[    ]";
    const opened = open(`<w:p>${control(run(placeholder))}${SLOT}</w:p>`);
    const paragraph = opened.state.doc.child(0);
    const typed = type(
      select(opened.state, 1, 1 + paragraph.content.size),
      "12"
    );

    expect(controlTexts(typed.doc)).toEqual(["12"]);
  });
});

describe("a blank inside a hyperlink", () => {
  it("gives its formatting to none of the text typed beside it", () => {
    const opened = open(
      `<w:p>${run("See ")}<w:hyperlink w:anchor="here">${SLOT}</w:hyperlink></w:p>`
    );
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toHaveLength(1);
    expect(highlightOf(textNode(typed.doc, "K"))).toBeUndefined();
  });

  it("keeps its blank, and the link goes out around it", () => {
    const opened = open(
      `<w:p><w:hyperlink w:anchor="here">${SLOT}</w:hyperlink></w:p>`
    );
    const typed = type(afterSlot(opened.state), "K");

    expect(emptyRuns(typed.doc)).toHaveLength(1);
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      `<w:hyperlink w:anchor="here">${SLOT.replace("></w:t>", "/>")}</w:hyperlink>`
    );
  });
});

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function openEditor(body: string): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const view = createEditorView({
    mount,
    state: open(body).state,
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

describe("a composition opening at a blank", () => {
  it("fills it, and the next syllable carries on in the control", () => {
    const view = openEditor(IN_CONTROL());
    const pos = firstEmptyRun(view.state.doc) + 1;
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, pos))
    );

    openComposition(view);
    view.dispatch(view.state.tr.insertText("가"));
    expect(emptyRuns(view.state.doc)).toEqual([]);
    expect(controlTexts(view.state.doc)).toEqual(["가"]);

    openComposition(view);
    view.dispatch(view.state.tr.insertText("나"));
    expect(controlTexts(view.state.doc)).toEqual(["가나"]);
  });
});

describe("a press on a blank", () => {
  it("puts the caret after it", () => {
    const view = openEditor(`<w:p>${run("Label: ")}${SLOT}</w:p>`);
    const pos = firstEmptyRun(view.state.doc);
    const node = view.state.doc.nodeAt(pos);
    if (!node) throw new Error("no blank");
    const handled = view.someProp("handleClickOn", (handler) =>
      handler(view, pos, node, pos, new MouseEvent("click"), true)
    );

    expect(handled).toBe(true);
    expect(view.state.selection.from).toBe(pos + 1);
  });
});

const TRACE = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';

const boldRun = (text: string) =>
  `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

function boldOf(node: PMNode | null | undefined): unknown {
  const run = node?.marks.find((mark) => mark.type === docxSchema.marks.run);
  const format: unknown = run?.attrs.format;
  return typeof format === "object" && format !== null && "bold" in format
    ? format.bold
    : undefined;
}

describe("a blank between markers that draw nothing", () => {
  const MARKED = [
    [
      "bookmark",
      `<w:p><w:bookmarkStart w:id="1" w:name="f"/>${SLOT}<w:bookmarkEnd w:id="1"/></w:p>`,
    ],
    [
      "proofing mark",
      `<w:p><w:proofErr w:type="spellStart"/>${SLOT}<w:proofErr w:type="spellEnd"/></w:p>`,
    ],
  ];

  it.each(
    MARKED.flatMap(([name, body]) => [0, 1, 2, 3].map((at) => [name, at, body]))
  )(
    "is filled by text typed beside a %s, %i steps into the paragraph",
    (_, at, body) => {
      const opened = open(String(body));
      const pos = 1 + Number(at);
      const typed = type(select(opened.state, pos, pos), "K");

      expect(emptyRuns(typed.doc)).toEqual([]);
      expect(highlightOf(textNode(typed.doc, "K"))).toBe("yellow");
    }
  );
});

describe("a run holding no characters that draws nothing", () => {
  it("is passed over by text typed after it, which carries on the word before it", () => {
    const opened = open(`<w:p>${boldRun("Text")}${TRACE}</w:p>`);
    const end = 1 + opened.state.doc.child(0).content.size;
    const typed = type(select(opened.state, end, end), "X");

    expect(boldOf(textNode(typed.doc, "X"))).toBe(true);
    expect(emptyRuns(typed.doc)).toHaveLength(1);
  });

  it("is passed over by text typed before it, which takes the word after it", () => {
    const opened = open(`<w:p>${TRACE}${boldRun("Text")}</w:p>`);
    const typed = type(select(opened.state, 1, 1), "X");

    expect(boldOf(textNode(typed.doc, "X"))).toBe(true);
    expect(emptyRuns(typed.doc)).toHaveLength(1);
  });

  it("gives its formatting to text typed in a paragraph that shows nothing else", () => {
    const opened = open(`<w:p>${TRACE}</w:p>`);
    const typed = type(select(opened.state, 2, 2), "X");

    expect(emptyRuns(typed.doc)).toEqual([]);
    expect(documentXmlOf(typed.doc, opened.session)).toContain(
      '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t xml:space="preserve">X</w:t></w:r>'
    );
  });
});

describe("a word before a marker that draws nothing", () => {
  it("gives its formatting to text typed after a bookmark", () => {
    const opened = open(
      `<w:p>${boldRun("Text")}<w:bookmarkStart w:id="1" w:name="f"/></w:p>`
    );
    const end = 1 + opened.state.doc.child(0).content.size;
    const typed = type(select(opened.state, end, end), "X");

    expect(boldOf(textNode(typed.doc, "X"))).toBe(true);
  });

  it("gives its formatting to text typed after a comment's range marker", () => {
    const { doc } = importDocx(
      commentedDocx(
        `<w:p>${boldRun("Text")}<w:commentRangeStart w:id="0"/>${run("word")}` +
          '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
        [{ id: "0", text: "note" }]
      )
    );
    const state = createEditorState(doc);
    const pos = posOfText(state.doc, "word") - 1;
    const typed = type(select(state, pos, pos), "X");

    expect(boldOf(textNode(typed.doc, "X"))).toBe(true);
  });
});

/** Hands the view a paste carrying these clipboard types, the way a browser does */
function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (kind: string) => data[kind] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

describe("a paste at a blank", () => {
  it.each([
    ["bold HTML", { "text/html": "<b>Kim</b>", "text/plain": "Kim" }],
    ["plain HTML", { "text/html": "<span>Kim</span>", "text/plain": "Kim" }],
    ["plain text", { "text/plain": "Kim" }],
  ])("of %s fills it with text in the blank's formatting", (_, data) => {
    const opened = open(CELL);
    const view = openEditor(CELL);
    const pos = firstEmptyRun(view.state.doc) + 1;
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, pos))
    );
    paste(view, data);

    expect(view.state.doc.textContent).toContain("Kim");
    expect(emptyRuns(view.state.doc)).toEqual([]);
    expect(documentXmlOf(view.state.doc, opened.session)).toContain(
      `<w:tc><w:p>${filledRun("Kim")}</w:p></w:tc>`
    );
  });
});
