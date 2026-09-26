// @vitest-environment jsdom
/**
 * An inline content control kept through an edit of everything it holds.
 *
 * Each case edits through a state built the way the editor builds one, so the transaction this
 * plugin appends is judged by the guards exactly as it is on screen, and the export is read back
 * to see the control where the file needs it: its tag and its id around the new text.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { Fragment, Slice } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx } from "../../__testing__/docx";
import {
  openComposition,
  rangeOfText,
  select,
} from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { docxSchema } from "../../schema";
import { wrappersOf } from "../../schema/wrappers";
import { undo } from "../commands/historyCommands";
import { createEditorView, editorStateForSession } from "../createEditor";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const boldRun = (text: string) =>
  `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A run holding nothing but its properties, which a word processor leaves at a paragraph's end */
const EMPTY_RUN = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';

const PLACEHOLDER = "[    ]";

function control(
  content: string,
  { id = 5, tag = "PERIOD", props = "" } = {}
): string {
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/>` +
    `<w:id w:val="${id}"/>${props}<w:text/></w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

const lock = (val: string) => `<w:lock w:val="${val}"/>`;

/** Text, the control, then text again, in one paragraph */
const BETWEEN = (props = "") =>
  `<w:p>${run("From ")}${control(run(PLACEHOLDER), { props })}${run(" on")}</w:p>`;

/** The control as the whole of its paragraph, an empty run standing after it */
const ALONE = `<w:p>${control(run(PLACEHOLDER))}${EMPTY_RUN}</w:p>`;

interface Opened {
  state: EditorState;
  session: ReturnType<typeof importDocx>["session"];
}

function open(body: string): Opened {
  const opened = importDocx(makeDocx(body));
  return { state: editorStateForSession(opened), session: opened.session };
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

function emptyControls(doc: PMNode): PMNode[] {
  const found: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === "sdtEmptyInline") found.push(node);
    return true;
  });
  return found;
}

/** The stretch the placeholder covers, which is the whole of the control */
function placeholder(doc: PMNode): { from: number; to: number } {
  return rangeOfText(doc, PLACEHOLDER);
}

function overPlaceholder(state: EditorState): EditorState {
  const { from, to } = placeholder(state.doc);
  return select(state, from, to);
}

/** Types each character at the selection, one transaction each, the way a keyboard does */
function type(state: EditorState, text: string): EditorState {
  let next = state;
  for (const char of text) next = next.apply(next.tr.insertText(char));
  return next;
}

/** The exported `w:sdt` carrying this tag, whole */
function exportedControl(opened: Opened, doc: PMNode, tag = "PERIOD"): string {
  const xml = documentXmlOf(doc, opened.session);
  const found = xml
    .split("<w:sdt>")
    .slice(1)
    .map((part) => `<w:sdt>${part.slice(0, part.indexOf("</w:sdt>"))}</w:sdt>`)
    .find((part) => part.includes(`<w:tag w:val="${tag}"/>`));
  if (found === undefined) throw new Error(`no control tagged ${tag}`);
  return found;
}

describe("typing over everything a control holds", () => {
  it("writes the text into the control, which goes out with its tag and id", () => {
    const opened = open(BETWEEN());
    const typed = type(overPlaceholder(opened.state), "2026");

    expect(controlTexts(typed.doc)).toEqual(["2026"]);
    expect(typed.doc.textContent).toBe("From 2026 on");
    const exported = exportedControl(opened, typed.doc);
    expect(exported).toContain('<w:id w:val="5"/>');
    expect(exported).toContain(">2026</w:t>");
  });

  it("reads a selection made from the end back to the start the same way", () => {
    const opened = open(BETWEEN());
    const { from, to } = placeholder(opened.state.doc);
    const typed = type(select(opened.state, to, from), "12");

    expect(controlTexts(typed.doc)).toEqual(["12"]);
  });

  it("sets aside what draws nothing past the control, as a triple click selects it", () => {
    const opened = open(ALONE);
    const paragraph = opened.state.doc.child(0);
    const typed = type(
      select(opened.state, 1, 1 + paragraph.content.size),
      "12"
    );

    expect(controlTexts(typed.doc)).toEqual(["12"]);
    expect(exportedControl(opened, typed.doc)).toContain(">12</w:t>");
  });

  it("keeps the text typed over the end of the control inside it", () => {
    const opened = open(BETWEEN());
    const { from, to } = placeholder(opened.state.doc);
    const typed = type(select(opened.state, from + 1, to), "1]");

    expect(controlTexts(typed.doc)).toEqual(["[1]"]);
  });

  it("keeps the formatting the replaced text wore", () => {
    const opened = open(
      `<w:p>${control(boldRun(PLACEHOLDER))}${run(" after")}</w:p>`
    );
    const typed = type(overPlaceholder(opened.state), "12");

    expect(exportedControl(opened, typed.doc)).toContain("<w:b/>");
  });

  it("types beside a control the caret was put against rather than into it", () => {
    const opened = open(BETWEEN());
    const typed = type(
      select(opened.state, placeholder(opened.state.doc).to),
      "x"
    );

    expect(controlTexts(typed.doc)).toEqual([PLACEHOLDER]);
    expect(typed.doc.textContent).toBe(`From ${PLACEHOLDER}x on`);
  });

  it("stops writing into the control once the caret moves", () => {
    const opened = open(BETWEEN());
    const typed = type(overPlaceholder(opened.state), "1");
    const end = typed.selection.from;
    const moved = select(select(typed, end - 1), end);

    expect(controlTexts(type(moved, "x").doc)).toEqual(["1"]);
  });

  it("is taken back by one undo", () => {
    const opened = open(BETWEEN());
    const typed = type(overPlaceholder(opened.state), "1");

    let undone: EditorState | null = null;
    undo(typed, (tr) => {
      undone = typed.apply(tr);
    });
    expect(undone).not.toBeNull();
    expect((undone as EditorState | null)?.doc.eq(opened.state.doc)).toBe(true);
  });
});

describe("pasting over everything a control holds", () => {
  it("puts the pasted text inside the control", () => {
    const opened = open(BETWEEN());
    const state = overPlaceholder(opened.state);
    const pasted = state.apply(
      state.tr.replaceSelection(
        new Slice(Fragment.from(docxSchema.text("2026-10-01")), 0, 0)
      )
    );

    expect(controlTexts(pasted.doc)).toEqual(["2026-10-01"]);
  });
});

describe("deleting everything a control holds", () => {
  it("leaves the control standing with nothing inside it", () => {
    const opened = open(BETWEEN());
    const state = overPlaceholder(opened.state);
    const deleted = state.apply(state.tr.deleteSelection());

    expect(deleted.doc.textContent).toBe("From  on");
    expect(emptyControls(deleted.doc)).toHaveLength(1);
    const exported = exportedControl(opened, deleted.doc);
    expect(exported).toContain('<w:id w:val="5"/>');
    expect(exported).not.toMatch(/<w:t[ >]/);
  });

  it("does the same where the last character goes on its own", () => {
    const opened = open(
      `<w:p>${run("From ")}${control(run("1"))}${run(" on")}</w:p>`
    );
    const at = rangeOfText(opened.state.doc, "1").from;
    const state = select(opened.state, at + 1);
    const deleted = state.apply(state.tr.delete(at, at + 1));

    expect(emptyControls(deleted.doc)).toHaveLength(1);
  });

  it("writes what is typed next back into the control, formatted as before", () => {
    const opened = open(
      `<w:p>${run("From ")}${control(boldRun(PLACEHOLDER))}${run(" on")}</w:p>`
    );
    const state = overPlaceholder(opened.state);
    const typed = type(state.apply(state.tr.deleteSelection()), "12");

    expect(emptyControls(typed.doc)).toHaveLength(0);
    expect(controlTexts(typed.doc)).toEqual(["12"]);
    expect(exportedControl(opened, typed.doc)).toContain("<w:b/>");
  });

  it("moves the control along with its text when the text is dragged away", () => {
    const opened = open(BETWEEN());
    const { from, to } = placeholder(opened.state.doc);
    const slice = opened.state.doc.slice(from, to);
    const tr = opened.state.tr.delete(from, to);
    tr.insert(tr.mapping.map(1), slice.content);
    const moved = opened.state.apply(tr);

    expect(emptyControls(moved.doc)).toHaveLength(0);
    expect(controlTexts(moved.doc)).toEqual([PLACEHOLDER]);
  });
});

describe("a selection reaching past the control", () => {
  it("takes the control away with the text beside it", () => {
    const opened = open(BETWEEN());
    const { from, to } = placeholder(opened.state.doc);
    const typed = type(select(opened.state, from - 1, to + 1), "x");

    expect(controlTexts(typed.doc)).toEqual([]);
    expect(emptyControls(typed.doc)).toHaveLength(0);
    expect(typed.doc.textContent).toBe("Fromxon");
  });

  it("is refused where the control is locked against deletion", () => {
    const opened = open(BETWEEN(lock("sdtLocked")));
    const { from, to } = placeholder(opened.state.doc);
    const typed = type(select(opened.state, from - 1, to + 1), "x");

    expect(typed.doc.eq(opened.state.doc)).toBe(true);
  });
});

/**
 * What each `w:lock` value makes of writing over everything the control holds: its contents where
 * they may be edited, the control itself where they may not (`schema/locks`).
 */
const LOCKS = [
  { val: "sdtLocked", written: true, removed: false },
  { val: "contentLocked", written: false, removed: true },
  { val: "sdtContentLocked", written: false, removed: false },
] as const;

describe.each(LOCKS)(
  "a control whose lock reads $val",
  ({ val, written, removed }) => {
    it(`takes the text typed over it: ${written}`, () => {
      const opened = open(BETWEEN(lock(val)));
      const typed = type(overPlaceholder(opened.state), "12");

      expect(controlTexts(typed.doc)).toEqual(
        written ? ["12"] : removed ? [] : [PLACEHOLDER]
      );
    });

    it(`is emptied rather than removed by a deletion: ${written}`, () => {
      const opened = open(BETWEEN(lock(val)));
      const state = overPlaceholder(opened.state);
      const deleted = state.apply(state.tr.deleteSelection());

      expect(emptyControls(deleted.doc)).toHaveLength(written ? 1 : 0);
      expect(controlTexts(deleted.doc)).toEqual(
        written || removed ? [] : [PLACEHOLDER]
      );
    });
  }
);

describe("what the control says about outliving an edit", () => {
  it("drops the placeholder flag once the text is written over", () => {
    const opened = open(BETWEEN("<w:showingPlcHdr/>"));
    const typed = type(overPlaceholder(opened.state), "12");

    const exported = exportedControl(opened, typed.doc);
    expect(exported).toContain(">12</w:t>");
    expect(exported).not.toContain("showingPlcHdr");
  });

  it("lets a temporary control go with the edit, written over or emptied", () => {
    const opened = open(BETWEEN("<w:temporary/>"));
    const state = overPlaceholder(opened.state);
    const typed = type(state, "12");
    const deleted = state.apply(state.tr.deleteSelection());

    expect(controlTexts(typed.doc)).toEqual([]);
    expect(typed.doc.textContent).toBe("From 12 on");
    expect(controlTexts(deleted.doc)).toEqual([]);
    expect(emptyControls(deleted.doc)).toHaveLength(0);
  });
});

describe("a control standing inside another", () => {
  const DATE = control(run(PLACEHOLDER), { id: 6, tag: "DATE" });

  it("keeps both where the inner one is written over, and the outer holds text beside it", () => {
    const opened = open(
      `<w:p>${control(run("On ") + DATE, { tag: "CLAUSE" })}</w:p>`
    );
    const typed = type(overPlaceholder(opened.state), "12");

    expect(controlTexts(typed.doc).sort()).toEqual(["12", "On 12"]);
    expect(exportedControl(opened, typed.doc, "DATE")).toContain(">12</w:t>");
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

describe("a composition opening over everything a control holds", () => {
  it("opens inside the control, which the text it composes then fills", () => {
    const view = openEditor(BETWEEN());
    const { from, to } = placeholder(view.state.doc);
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
    );

    openComposition(view);
    // The composition opens with the control among the marks it writes with
    const stored = view.state.storedMarks ?? [];
    expect(stored.some((mark) => mark.type === docxSchema.marks.sdt)).toBe(
      true
    );

    view.dispatch(view.state.tr.insertText("가"));
    expect(controlTexts(view.state.doc)).toEqual(["가"]);
    expect(emptyControls(view.state.doc)).toHaveLength(0);
  });

  it("writes the next syllable into the control the caret is still writing into", () => {
    const view = openEditor(BETWEEN());
    const { from, to } = placeholder(view.state.doc);
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
    );
    view.dispatch(view.state.tr.insertText("가"));
    // No marks are stored for the caret a syllable leaves, so the control is not among them
    expect(view.state.storedMarks).toBeNull();

    openComposition(view);
    view.dispatch(view.state.tr.insertText("나"));
    expect(controlTexts(view.state.doc)).toEqual(["가나"]);
  });
});
