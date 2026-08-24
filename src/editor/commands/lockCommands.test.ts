// @vitest-environment jsdom
import type { Mark, Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { makeDocx, makeStyledDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { buildParagraph } from "../../docx/importParagraph";
import { serializeParagraph } from "../../docx/serializeParagraph";
import { serializeTable } from "../../docx/serializeTable";
import { parseXml, W_NS } from "../../ooxml/xml";
import { addRowAfter } from "../../table";
import { createEditorState } from "../createEditor";
import { insertTable } from "../insertTable";
import { canRunCommand } from "./canRunCommand";
// The history commands as a consumer reaches them, since a raw prosemirror-history one would meet
// the guard's refusal
import { redo, undo } from "./index";
import {
  documentHasLocked,
  lockSelection,
  selectionLock,
  selectionTouchesLocked,
  unlockSelection,
} from "./lockCommands";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const control = (inner: string, pr = LOCKED_PR) =>
  `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const sdt = (inner: string) => control(inner);

const lockPr = (val: string) =>
  `<w:sdtPr><w:id w:val="7"/><w:lock w:val="${val}"/></w:sdtPr>`;

/** The XML a newly made lock writes, which is what Word reads back as a shut control */
const NEW_CONTROL =
  /^<w:sdt><w:sdtPr><w:id w:val="\d+"\/><w:lock w:val="sdtContentLocked"\/><\/w:sdtPr>$/;

function opened(body: string): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc);
}

function selected(state: EditorState, from: number, to: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

/** Runs the command over the selection the state already holds */
function ranHere(state: EditorState, command: Command): EditorState {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** Runs the command over this selection and hands back the state it leaves */
function ran(
  state: EditorState,
  command: Command,
  from: number,
  to = from
): EditorState {
  return ranHere(selected(state, from, to), command);
}

function sdtMarkOf(node: PMNode): Mark {
  const mark = node.marks.find((entry) => entry.type.name === "sdt");
  if (!mark) throw new Error("no sdt mark");
  return mark;
}

/** The controls one paragraph holds, one entry per inline node that wears one */
function controlsOf(paragraph: PMNode): Mark[] {
  const marks: Mark[] = [];
  paragraph.forEach((child) => {
    const mark = child.marks.find((entry) => entry.type.name === "sdt");
    if (mark) marks.push(mark);
  });
  return marks;
}

describe("locking the selected text", () => {
  it("wraps it in a control that carries an id and the lock", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    const next = ran(state, lockSelection, 1, 4);

    const mark = sdtMarkOf(next.doc.child(0).child(0));
    // The value written shuts both clauses of the lock
    expect(mark.attrs.contentsLocked).toBe(true);
    expect(mark.attrs.deletionLocked).toBe(true);
    expect(mark.attrs.sdtPrefix).toMatch(NEW_CONTROL);
    expect(next.doc.child(0).textContent).toBe("abc");
  });

  it("makes a control of its own per paragraph the selection runs through", () => {
    const state = opened(`<w:p>${run("ab")}</w:p><w:p>${run("cd")}</w:p>`);
    const next = ran(state, lockSelection, 1, 7);

    const first = controlsOf(next.doc.child(0));
    const second = controlsOf(next.doc.child(1));
    expect(first).toHaveLength(1);
    expect(second).toHaveLength(1);
    // A control cannot cross a paragraph, so these are two controls with ids of their own
    expect(first[0].attrs.sdtPrefix).toMatch(NEW_CONTROL);
    expect(second[0].attrs.sdtPrefix).toMatch(NEW_CONTROL);
    expect(first[0].attrs.sdtPrefix).not.toBe(second[0].attrs.sdtPrefix);
  });

  it("has nothing to do where a control already stands", () => {
    const state = opened(`<w:p>${sdt(run("bc"))}</w:p>`);
    expect(lockSelection(selected(state, 1, 3))).toBe(false);
  });

  it("locks only the paragraphs that hold no control yet", () => {
    const state = opened(`<w:p>${sdt(run("bc"))}</w:p><w:p>${run("ab")}</w:p>`);
    const next = ran(state, lockSelection, 1, 7);

    expect(controlsOf(next.doc.child(0))[0].attrs.sdtPrefix).toBe(
      `<w:sdt>${LOCKED_PR}`
    );
    expect(controlsOf(next.doc.child(1))[0].attrs.sdtPrefix).toMatch(
      NEW_CONTROL
    );
  });

  it("has nothing to do with no text selected", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    expect(lockSelection(selected(state, 2, 2))).toBe(false);
  });
});

describe("lifting a lock", () => {
  /** Plain text, a locked control over two runs, then plain text again */
  const BODY = `<w:p>${run("a")}${sdt(run("b") + run("c"))}${run("d")}</w:p>`;

  it("opens the whole stretch of the control from a caret inside it", () => {
    const state = opened(BODY);
    const next = ran(state, unlockSelection, 3);

    expect(documentHasLocked(next.doc)).toBe(false);
    expect(next.doc.child(0).textContent).toBe("abcd");
    // The control stays where it stood, and only its lock is gone
    const controls = controlsOf(next.doc.child(0));
    expect(controls.map((mark) => mark.attrs.contentsLocked)).toEqual([false]);
    expect(controls.map((mark) => mark.attrs.deletionLocked)).toEqual([false]);
    expect(controls[0].attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
  });

  it("reaches the control from a caret resting against its edge", () => {
    const state = opened(BODY);
    expect(unlockSelection(selected(state, 2, 2))).toBe(true);
    expect(unlockSelection(selected(state, 4, 4))).toBe(true);
  });

  it("has nothing to do where the selection reaches no locked control", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    expect(unlockSelection(selected(state, 1, 4))).toBe(false);
  });

  it("reports the caret as shut only inside the control, not against its edges", () => {
    const state = opened(BODY);
    expect(selectionTouchesLocked(selected(state, 3, 3))).toBe(true);
    expect(selectionTouchesLocked(selected(state, 2, 2))).toBe(false);
    expect(selectionTouchesLocked(selected(state, 4, 4))).toBe(false);
    // A selection that runs over the locked text is shut whichever way it was made
    expect(selectionTouchesLocked(selected(state, 1, 5))).toBe(true);
  });
});

/**
 * The one answer the commands and the menus are built on, so what it names is what the menu draws
 */
describe("the lock state of the selection", () => {
  /** A paragraph of plain text, a locked control and plain text again, then a plain paragraph */
  const BODY =
    `<w:p>${run("a")}${sdt(run("bc"))}${run("d")}</w:p>` +
    `<w:p>${run("ef")}</w:p>`;

  it("is none where the selection reaches neither a lock nor text to lock", () => {
    expect(selectionLock(selected(opened(BODY), 8, 8))).toBe("none");
  });

  it("is lockable over text no control stands in", () => {
    expect(selectionLock(selected(opened(BODY), 7, 9))).toBe("lockable");
  });

  it("is locked over a stretch a control already shuts", () => {
    expect(selectionLock(selected(opened(BODY), 2, 4))).toBe("locked");
  });

  it("is mixed where the selection holds both, and both commands still run", () => {
    const over = selected(opened(BODY), 1, 5);

    expect(selectionLock(over)).toBe("mixed");
    // Which of the two the menu offers is the menu's own policy, not this function's
    expect(lockSelection(over)).toBe(true);
    expect(unlockSelection(over)).toBe(true);
  });

  it("is locked over a cell a control shuts as a whole", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    expect(selectionLock(selected(state, inside, inside))).toBe("locked");
  });
});

describe("the locks a document holds", () => {
  it("are found in an imported document and gone once lifted", () => {
    const state = opened(`<w:p>${run("a")}${sdt(run("bc"))}</w:p>`);
    expect(documentHasLocked(state.doc)).toBe(true);
    expect(documentHasLocked(ran(state, unlockSelection, 3).doc)).toBe(false);
  });

  it("are none in a document nobody locked", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    expect(documentHasLocked(state.doc)).toBe(false);
    expect(documentHasLocked(ran(state, lockSelection, 1, 4).doc)).toBe(true);
  });
});

/**
 * What the three public answers say about each of the four values `w:lock` takes, now that a
 * control can be locked against deletion while its contents stand open.
 *
 * `selectionTouchesLocked` and `selectionLock` are about editing where the selection stands, so
 * both read the contents clause: a control locked against deletion alone is text to lock rather
 * than a lock in reach. `documentHasLocked` is about the document instead, so either clause counts.
 */
describe.each([
  { val: "unlocked", shut: false, lock: "lockable", held: false },
  { val: "sdtLocked", shut: false, lock: "lockable", held: true },
  { val: "contentLocked", shut: true, lock: "locked", held: true },
  { val: "sdtContentLocked", shut: true, lock: "locked", held: true },
])("a control whose lock reads $val", ({ val, shut, lock, held }) => {
  const BODY = `<w:p>${run("a")}${control(run("bc"), lockPr(val))}${run("d")}</w:p>`;

  it(`shuts editing where the caret stands inside it: ${shut}`, () => {
    expect(selectionTouchesLocked(selected(opened(BODY), 3, 3))).toBe(shut);
  });

  it(`reads as ${lock} over the text it holds`, () => {
    expect(selectionLock(selected(opened(BODY), 2, 4))).toBe(lock);
  });

  it(`counts as a lock the document holds: ${held}`, () => {
    expect(documentHasLocked(opened(BODY).doc)).toBe(held);
  });
});

/**
 * The same three answers about a cell, whose control carries the two clauses as attributes of the
 * cell itself.
 */
describe.each([
  { val: "unlocked", shut: false, held: false },
  { val: "sdtLocked", shut: false, held: true },
  { val: "contentLocked", shut: true, held: true },
  { val: "sdtContentLocked", shut: true, held: true },
])("a cell inside a control whose lock reads $val", ({ val, shut, held }) => {
  const BODY =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    `<w:tr>${wrappedCell(lockPr(val), "locked cell")}` +
    `<w:tc><w:p>${run("next cell")}</w:p></w:tc></w:tr>` +
    "</w:tbl>";

  it(`shuts editing where the caret stands inside it: ${shut}`, () => {
    const state = opened(BODY);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    expect(selectionTouchesLocked(selected(state, inside, inside))).toBe(shut);
  });

  it(`counts as a lock the document holds: ${held}`, () => {
    expect(documentHasLocked(opened(BODY).doc)).toBe(held);
  });
});

function undone(state: EditorState): EditorState {
  let next = state;
  undo(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

function redone(state: EditorState): EditorState {
  let next = state;
  redo(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** The opening XML of the one control the paragraph holds */
function controlPrefix(paragraph: PMNode): unknown {
  const controls = controlsOf(paragraph);
  expect(controls).toHaveLength(1);
  return controls[0].attrs.sdtPrefix;
}

describe("the undo history around a lock", () => {
  it("takes a lock back, leaving the text it shut open to editing again", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    const back = undone(ran(state, lockSelection, 1, 4));

    expect(documentHasLocked(back.doc)).toBe(false);
    expect(controlsOf(back.doc.child(0))).toEqual([]);
    const typed = back.apply(back.tr.insertText("x", 2));
    expect(typed.doc.child(0).textContent).toBe("axbc");
  });

  it("shuts the lock again where a redo asks for it", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    const back = undone(ran(state, lockSelection, 1, 4));
    expect(documentHasLocked(back.doc)).toBe(false);

    const again = redone(back);
    expect(documentHasLocked(again.doc)).toBe(true);
    expect(controlPrefix(again.doc.child(0))).toMatch(NEW_CONTROL);
  });

  it("takes a lifted lock back, on the very control it came off", () => {
    const state = opened(`<w:p>${run("a")}${sdt(run("bc"))}</w:p>`);
    const back = undone(ran(state, unlockSelection, 3));

    expect(documentHasLocked(back.doc)).toBe(true);
    expect(controlPrefix(back.doc.child(0))).toBe(`<w:sdt>${LOCKED_PR}`);
  });

  it("lifts the lock again where a redo asks for it", () => {
    const state = opened(`<w:p>${run("a")}${sdt(run("bc"))}</w:p>`);
    const back = undone(ran(state, unlockSelection, 3));
    expect(documentHasLocked(back.doc)).toBe(true);

    const again = redone(back);
    expect(documentHasLocked(again.doc)).toBe(false);
    expect(controlPrefix(again.doc.child(0))).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
  });

  /** The redo of a cell unlock is the step that puts the cell's own lock down again */
  it("takes the lock of a cell back and forth", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    const lifted = ran(state, unlockSelection, inside);
    const cellLock = (at: EditorState) =>
      cellOf(at.doc, "locked cell").node.attrs.sdtContentsLocked;

    expect(cellLock(lifted)).toBe(false);
    expect(cellLock(undone(lifted))).toBe(true);
    expect(cellLock(redone(undone(lifted)))).toBe(false);
  });

  it("gives back the lock, then what was typed before it, one step each", () => {
    const state = opened(`<w:p>${run("ab")}</w:p>`);
    const typed = state.apply(state.tr.insertText("c", 3));
    const locked = ran(typed, lockSelection, 1, 3);

    const once = undone(locked);
    expect(documentHasLocked(once.doc)).toBe(false);
    expect(once.doc.child(0).textContent).toBe("abc");

    expect(undone(once).doc.child(0).textContent).toBe("ab");
  });

  it("leaves a lock the document came in with, having no event that shut it", () => {
    const state = opened(`<w:p>${run("a")}${sdt(run("bc"))}</w:p>`);
    const typed = state.apply(state.tr.insertText("d", 2));
    const back = undone(typed);

    expect(back.doc.child(0).textContent).toBe("abc");
    expect(documentHasLocked(back.doc)).toBe(true);
    // Nothing in this session shut that control, so the history holds nothing that would open it
    expect(undo(back)).toBe(false);
  });
});

const lockedCell = (text: string) =>
  `<w:tc><w:p>${sdt(run(text))}</w:p></w:tc>`;

/** A two by two table with a locked control in every cell */
const LOCKED_TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${lockedCell("TopLeft")}${lockedCell("TopRight")}</w:tr>` +
  `<w:tr>${lockedCell("BottomLeft")}${lockedCell("BottomRight")}</w:tr>` +
  "</w:tbl>";

/** The position just before each cell, which is what a cell selection is built from */
function cellPositions(doc: PMNode): number[] {
  const found: number[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") found.push(pos);
    return true;
  });
  return found;
}

function lockedTexts(doc: PMNode): string[] {
  const texts: string[] = [];
  doc.descendants((node) => {
    const locked = node.marks.some(
      (mark) => mark.type.name === "sdt" && mark.attrs.contentsLocked === true
    );
    if (node.isText && locked) texts.push(node.text ?? "");
    return true;
  });
  return texts;
}

describe("lifting locks from a block of table cells", () => {
  /** The left column runs from the first cell to the third, taking in an unselected one between */
  it("reaches only the controls standing in the selected cells", () => {
    const state = opened(LOCKED_TABLE);
    const cells = cellPositions(state.doc);
    const column = state.apply(
      state.tr.setSelection(CellSelection.create(state.doc, cells[0], cells[2]))
    );

    let next = column;
    expect(
      unlockSelection(column, (tr) => {
        next = column.apply(tr);
      })
    ).toBe(true);
    expect(lockedTexts(next.doc)).toEqual(["TopRight", "BottomRight"]);
  });
});

/** The single element a fragment holds */
function element(xml: string): Element {
  const wrapped = parseXml(`<w:wrap xmlns:w="${W_NS}">${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return el;
}

describe("a lock made in the editor", () => {
  it("is written out as a Word control and comes back locked", () => {
    const state = opened(`<w:p>${run("abc")}</w:p>`);
    const locked = ran(state, lockSelection, 1, 4);
    const written = serializeParagraph(locked.doc.child(0));

    expect(written).toContain('<w:lock w:val="sdtContentLocked"/>');
    expect(written).toContain("<w:sdtContent>");

    const reopened = buildParagraph(element(written), 0);
    if (!reopened) throw new Error("the paragraph could not be modelled");
    expect(reopened.textContent).toBe("abc");
    expect(sdtMarkOf(reopened.child(0))).toEqual(
      sdtMarkOf(locked.doc.child(0).child(0))
    );
    expect(documentHasLocked(reopened)).toBe(true);
  });
});

describe("a control Word wrote, once its lock is lifted", () => {
  const WORD_PR =
    "<w:sdtPr>" +
    '<w:alias w:val="signedOn"/><w:tag w:val="date"/><w:id w:val="7"/>' +
    '<w:lock w:val="sdtContentLocked"/><w:dataBinding w:xpath="/root/date"/>' +
    "</w:sdtPr>";

  /** Four characters of plain text, then the control over ten more */
  const BODY =
    "<w:p>" +
    run("on: ") +
    `<w:sdt>${WORD_PR}<w:sdtContent>${run("2026-08-04")}</w:sdtContent></w:sdt>` +
    "</w:p>";

  const INSIDE = 6;
  const written = (state: EditorState) =>
    serializeParagraph(state.doc.child(0));

  it("goes back out as the same control with nothing but the lock gone", () => {
    const xml = written(ran(opened(BODY), unlockSelection, INSIDE));

    expect(xml).toContain('<w:alias w:val="signedOn"/>');
    expect(xml).toContain('<w:tag w:val="date"/>');
    expect(xml).toContain('<w:id w:val="7"/>');
    expect(xml).toContain('<w:dataBinding w:xpath="/root/date"/>');
    expect(xml).toContain("<w:sdtContent>");
    expect(xml).not.toContain("<w:lock");
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
  });

  it("is shut again as the very same control, not wrapped in a second one", () => {
    const lifted = ran(opened(BODY), unlockSelection, INSIDE);
    // Part of the text selected again, the way a user would drag over it
    const shut = ran(lifted, lockSelection, INSIDE, INSIDE + 2);
    const xml = written(shut);

    expect(
      controlsOf(shut.doc.child(0)).map((mark) => mark.attrs.contentsLocked)
    ).toEqual([true]);
    expect(shut.doc.child(0).textContent).toBe("on: 2026-08-04");
    expect(xml).toContain('<w:id w:val="7"/>');
    expect(xml).toContain('<w:lock w:val="sdtContentLocked"/>');
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
  });

  it("takes a control of its own for the plain text a wider selection reaches", () => {
    const lifted = ran(opened(BODY), unlockSelection, INSIDE);
    const shut = ran(lifted, lockSelection, 1, INSIDE);
    const controls = controlsOf(shut.doc.child(0));

    expect(controls.map((mark) => mark.attrs.contentsLocked)).toEqual([
      true,
      true,
    ]);
    // The control Word wrote kept its own name, and the text beside it got one of its own
    expect(controls[0].attrs.sdtPrefix).toMatch(NEW_CONTROL);
    expect(controls[1].attrs.sdtPrefix).toContain('<w:id w:val="7"/>');
    expect(written(shut).match(/<w:sdt>/g)).toHaveLength(2);
  });
});

describe("a lock this editor made, once it is lifted", () => {
  const BODY = `<w:p>${run("before ")}${control(run("middle"))}${run(" after")}</w:p>`;
  const whole = (state: EditorState) => state.doc.child(0).content.size + 1;

  it("takes the paragraph over as one control when it is locked again", () => {
    const start = opened(BODY);
    const lifted = ran(start, unlockSelection, 1, whole(start));
    const shut = ran(lifted, lockSelection, 1, whole(lifted));
    const paragraph = shut.doc.child(0);

    expect(
      serializeParagraph(paragraph).match(/<w:sdt>/g),
      "the paragraph came back as more than one control"
    ).toHaveLength(1);
    expect(serializeParagraph(paragraph)).toContain('<w:id w:val="7"/>');
    expect(paragraph.textContent).toBe("before middle after");
    // One control over the whole of it leaves the three runs wearing one mark, so they are one run
    expect(
      controlsOf(paragraph).map((mark) => mark.attrs.contentsLocked)
    ).toEqual([true]);
    expect(paragraph.childCount).toBe(1);
  });
});

const wrappedCell = (pr: string, text: string) =>
  `<w:sdt>${pr}<w:sdtContent><w:tc><w:p>${run(text)}</w:p></w:tc>` +
  "</w:sdtContent></w:sdt>";

/** A one by two table whose left cell stands inside a control that shuts it */
const LOCKED_CELL_TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${wrappedCell(LOCKED_PR, "locked cell")}` +
  `<w:tc><w:p>${run("next cell")}</w:p></w:tc></w:tr>` +
  "</w:tbl>";

/** The cell holding this text, and where it stands */
function cellOf(doc: PMNode, needle: string): { pos: number; node: PMNode } {
  let found: { pos: number; node: PMNode } | null = null;
  doc.descendants((node, pos) => {
    if (
      !found &&
      node.type.name === "tableCell" &&
      node.textContent === needle
    ) {
      found = { pos, node };
    }
    return true;
  });
  if (!found) throw new Error(`no cell holding: ${needle}`);
  return found;
}

/** The same table, with a control of its own standing open inside the shut cell */
const OPEN_CONTROL_IN_LOCKED_CELL =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  "<w:tr>" +
  `<w:sdt>${LOCKED_PR}<w:sdtContent><w:tc><w:p>` +
  '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr>' +
  `<w:sdtContent>${run("inner")}</w:sdtContent></w:sdt>` +
  "</w:p></w:tc></w:sdtContent></w:sdt>" +
  `<w:tc><w:p>${run("next cell")}</w:p></w:tc>` +
  "</w:tr></w:tbl>";

describe("a cell Word locked", () => {
  it("is offered no locking of its own while it is shut", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    expect(lockSelection(selected(state, inside, inside + 2))).toBe(false);
  });

  /** A control cannot be shut where the guard would refuse the step that shuts it */
  it("is offered no shutting of the open control it holds either", () => {
    const state = opened(OPEN_CONTROL_IN_LOCKED_CELL);
    const inside = cellOf(state.doc, "inner").pos + 2;
    const over = selected(state, inside, inside + 2);

    expect(canRunCommand(lockSelection, over)).toBe(false);
    expect(lockSelection(over)).toBe(false);
  });

  it("is unlocked as a whole from a caret inside it, control and all", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    expect(unlockSelection(selected(state, inside, inside))).toBe(true);

    const next = ran(state, unlockSelection, inside);
    const cell = cellOf(next.doc, "locked cell").node;
    expect(cell.attrs.sdtContentsLocked).toBe(false);
    expect(cell.attrs.sdtDeletionLocked).toBe(false);
    expect(cell.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
    expect(documentHasLocked(next.doc)).toBe(false);
    // The control is still the one Word wrote, so it goes back out around the cell
    expect(serializeTable(next.doc.child(0))).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:tc>'
    );
  });

  it("takes an edit once the lock is off it", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = cellOf(state.doc, "locked cell").pos + 2;
    const lifted = ran(state, unlockSelection, inside);
    const typed = lifted.apply(lifted.tr.insertText("edit", inside));

    expect(typed.doc.textContent).toContain("edit");
  });
});

/** A default style, which is what gives `editor/styledParagraphs` something to lay on a paragraph */
const NORMAL_STYLE =
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
  '<w:name w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>';

/** The state `DocxEditor` builds, styles and all, so the plugins that append a transaction have work */
function openedStyled(body: string): EditorState {
  const { doc, session } = importDocx(makeStyledDocx(body, NORMAL_STYLE));
  return createEditorState(doc, {
    styles: session.styles,
    defaults: session.defaults,
    paragraphStyles: session.paragraphStyles,
  });
}

/**
 * `table/gridBorders` and `editor/styledParagraphs` each append a transaction of their own, which
 * carries no pass through the guard, so a replay has to leave them nothing the guard would refuse.
 */
describe("a lock among the structure edits of a table", () => {
  it("goes back and comes again whole, appended transactions and all", () => {
    const state = openedStyled(`<w:p>${run("body")}</w:p>`);
    const table = ranHere(state, insertTable({ rows: 2, columns: 2 }));
    const typed = table.apply(
      table.tr.insertText("cell", table.selection.head)
    );
    const at = cellOf(typed.doc, "cell").pos + 2;
    const locked = ran(typed, lockSelection, at, at + 4);
    const grown = ranHere(locked, addRowAfter);

    expect(documentHasLocked(grown.doc)).toBe(true);
    expect(grown.doc.child(1).childCount).toBe(3);

    const withoutRow = undone(grown);
    expect(withoutRow.doc.eq(locked.doc)).toBe(true);

    const withoutLock = undone(withoutRow);
    expect(documentHasLocked(withoutLock.doc)).toBe(false);
    expect(withoutLock.doc.eq(typed.doc)).toBe(true);

    expect(redone(redone(withoutLock)).doc.eq(grown.doc)).toBe(true);
  });
});
