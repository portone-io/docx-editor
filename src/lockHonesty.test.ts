// @vitest-environment jsdom
import { Mark, type Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { makeDocx, TINY_PNG_DATA_URL } from "./__testing__/docx";
import { select } from "./__testing__/editing";
import { importDocx } from "./docx/importDocx";
import * as commands from "./editor/commands/index";
import { createEditorState } from "./editor/createEditor";
import { setProtection } from "./editor/plugins/documentProtection";
import type { EditingProtection } from "./schema/protection";
import * as table from "./table";

/**
 * Every command the package exports answers truthfully over locked content, and under every
 * protection the editor can run under (`schema/protection`).
 *
 * The invariant: a command reporting false must change nothing when it is dispatched, and one
 * reporting true must change something. A command that reports true and is then refused by the
 * lock guard draws a live control that swallows the click; one that reports false and dispatches
 * anyway would be worse still. Both are the same fault, and this file is the only thing that
 * checks it, so the list of commands is written out by hand: half of them are factories
 * (`setParagraphAlign(align)`, `insertTable({rows, columns})`) that no reflection would reach.
 *
 * `NOT_A_COMMAND` carries everything else the two entries export, each with the reason it is not
 * a command, and the last test here fails the moment an export appears in neither list. That is
 * what stops a new command from being added without an answer to this question.
 */

const runXml = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const lockPr = (val: string) =>
  `<w:sdtPr><w:id w:val="7"/><w:lock w:val="${val}"/></w:sdtPr>`;

/** The value that shuts both clauses of the lock, which is the one this editor writes */
const lockedPr = lockPr("sdtContentLocked");

const control = (inner: string, pr = lockedPr) =>
  `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const cellXml = (inner: string) => `<w:tc><w:p>${inner}</w:p></w:tc>`;

/**
 * The shapes a lock comes in, in both the places it can stand and over each of the values whose
 * two clauses differ.
 *
 * A body paragraph holding a control that shuts both clauses over "bc", then one holding the two
 * controls with a single clause each: `contentLocked` over "ef", whose contents are shut and which
 * may still be taken away whole, and `sdtLocked` over "gh", whose contents stand open and which may
 * not. Then a two by two table whose top left cell stands inside a control that shuts both
 * clauses, and a three column table carrying the same two single-clause values on its cells.
 */
const BODY =
  `<w:p>${runXml("a")}${control(runXml("bc"))}${runXml("d")}</w:p>` +
  "<w:p>" +
  control(runXml("ef"), lockPr("contentLocked")) +
  control(runXml("gh"), lockPr("sdtLocked")) +
  "</w:p>" +
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${control(cellXml(runXml("TopLeft")))}${cellXml(runXml("TopRight"))}</w:tr>` +
  `<w:tr>${cellXml(runXml("BottomLeft"))}${cellXml(runXml("BottomRight"))}</w:tr>` +
  "</w:tbl>" +
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>' +
  '<w:gridCol w:w="1000"/></w:tblGrid>' +
  "<w:tr>" +
  control(cellXml(runXml("ShutCell")), lockPr("contentLocked")) +
  control(cellXml(runXml("KeptCell")), lockPr("sdtLocked")) +
  cellXml(runXml("PlainCell")) +
  "</w:tr>" +
  `<w:tr>${cellXml(runXml("Under1"))}${cellXml(runXml("Under2"))}${cellXml(runXml("Under3"))}</w:tr>` +
  "</w:tbl>";

function opened(protection: EditingProtection): EditorState {
  return createEditorState(importDocx(makeDocx(BODY)).doc, {
    protection,
    author: { id: "me", name: "Me" },
  });
}

/** The three standings a state can be built with, each of which every command is put to */
const PROTECTIONS: readonly EditingProtection[] = [
  "none",
  "comments",
  "readOnly",
];

const COMMENTED_RUN =
  '<w:commentRangeStart w:id="0"/>' +
  runXml("Commented") +
  '<w:commentRangeEnd w:id="0"/>' +
  '<w:r><w:commentReference w:id="0"/></w:r>';

function commentState(
  locked: boolean,
  protection: EditingProtection = "none"
): EditorState {
  const content = locked ? control(COMMENTED_RUN) : COMMENTED_RUN;
  return createEditorState(importDocx(makeDocx(`<w:p>${content}</w:p>`)).doc, {
    protection,
    author: { id: "me", name: "Me" },
  });
}

/** The position just inside the first text node reading exactly this */
function insideText(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text === needle) found = pos + 1;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return found;
}

/** The position of the cell holding this text */
function cellPos(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (
      found < 0 &&
      node.type.spec.tableRole === "cell" &&
      node.textContent === needle
    ) {
      found = pos;
    }
    return true;
  });
  if (found < 0) throw new Error(`no cell holding: ${needle}`);
  return found;
}

function caretIn(needle: string, protection: EditingProtection): EditorState {
  const state = opened(protection);
  return select(state, insideText(state.doc, needle));
}

/** The selection covering exactly this text, which is the whole of the control wrapping it */
function overText(needle: string, protection: EditingProtection): EditorState {
  const state = opened(protection);
  const inside = insideText(state.doc, needle);
  return select(state, inside - 1, inside - 1 + needle.length);
}

function cellsSelected(
  anchor: string,
  head: string,
  protection: EditingProtection
): EditorState {
  const state = opened(protection);
  return state.apply(
    state.tr.setSelection(
      CellSelection.create(
        state.doc,
        cellPos(state.doc, anchor),
        cellPos(state.doc, head)
      )
    )
  );
}

/** Where the selection stands, which is what decides whether a lock is in the way */
interface Place {
  name: string;
  state: (protection: EditingProtection) => EditorState;
}

const PLACES: readonly Place[] = [
  {
    name: "a caret in body text no control covers",
    state: (protection) => caretIn("a", protection),
  },
  {
    name: "a caret inside a locked control",
    state: (protection) => caretIn("bc", protection),
  },
  {
    name: "a selection running across a locked control",
    state: (protection) => {
      const state = opened(protection);
      return select(state, 1, state.doc.child(0).content.size + 1);
    },
  },
  {
    name: "a caret inside a locked cell",
    state: (protection) => caretIn("TopLeft", protection),
  },
  {
    name: "a caret in a cell no lock stands in",
    state: (protection) => caretIn("BottomRight", protection),
  },
  {
    name: "a block of cells one of which is locked",
    state: (protection) => cellsSelected("TopLeft", "TopRight", protection),
  },
  {
    name: "a block of cells no lock stands in",
    state: (protection) =>
      cellsSelected("BottomLeft", "BottomRight", protection),
  },
  {
    name: "a caret inside a control whose contents alone are locked",
    state: (protection) => caretIn("ef", protection),
  },
  {
    name: "a selection covering such a control whole, which may still be deleted",
    state: (protection) => overText("ef", protection),
  },
  {
    name: "a caret inside a control locked against deletion alone",
    state: (protection) => caretIn("gh", protection),
  },
  {
    name: "a selection covering a control locked against deletion alone whole",
    state: (protection) => overText("gh", protection),
  },
  {
    name: "a caret inside a cell whose contents alone are locked",
    state: (protection) => caretIn("ShutCell", protection),
  },
  {
    name: "a caret inside a cell locked against deletion alone",
    state: (protection) => caretIn("KeptCell", protection),
  },
  {
    name: "a block of cells one of which is locked against deletion alone",
    state: (protection) => cellsSelected("KeptCell", "PlainCell", protection),
  },
  {
    // The history is empty in a freshly opened document, so undo and redo have nothing to take
    // back anywhere else, and this is where their answering true is put to the test
    name: "a caret after an edit, so the history holds something",
    state: (protection) => {
      // The edit is made with nothing shut and the protection put on after, the way a mode is
      // switched on a document already worked on, so the history holds something under each
      const edited = caretIn("a", "none");
      const state = edited.apply(
        edited.tr.insertText("x", edited.selection.from)
      );
      return state.apply(
        setProtection(state.tr, {
          protection,
          author: { id: "me", name: "Me" },
          editableComments: "own",
        })
      );
    },
  },
];

const A_PICTURE = {
  src: TINY_PNG_DATA_URL,
  extent: { cx: 952500, cy: 952500 },
};

/** One exported command, under the name it is exported as */
interface CommandCase {
  /** The exported name, which is what the completeness test counts */
  name: string;
  command: Command;
}

const CASES: readonly CommandCase[] = [
  {
    name: "addComment",
    command: commands.addComment({ text: "note", author: "Author" }),
  },
  {
    name: "addCommentReply",
    command: commands.addCommentReply("0", {
      text: "reply",
      author: "Author",
    }),
  },
  { name: "updateComment", command: commands.updateComment("0", "note") },
  {
    name: "updateCommentReply",
    command: commands.updateCommentReply("0", "1", "reply"),
  },
  { name: "removeComment", command: commands.removeComment("0") },
  {
    name: "removeCommentReply",
    command: commands.removeCommentReply("0", "1"),
  },
  {
    name: "setCommentResolved",
    command: commands.setCommentResolved("0", true),
  },
  { name: "selectComment", command: commands.selectComment("0") },
  { name: "decreaseIndent", command: commands.decreaseIndent },
  { name: "increaseIndent", command: commands.increaseIndent },
  { name: "decreaseListLevel", command: commands.decreaseListLevel },
  { name: "increaseListLevel", command: commands.increaseListLevel },
  { name: "toggleBulletList", command: commands.toggleBulletList },
  { name: "toggleNumberedList", command: commands.toggleNumberedList },
  { name: "toggleBold", command: commands.toggleBold },
  { name: "toggleItalic", command: commands.toggleItalic },
  { name: "toggleStrike", command: commands.toggleStrike },
  { name: "toggleUnderline", command: commands.toggleUnderline },
  { name: "setFontFamily", command: commands.setFontFamily("Arial") },
  { name: "setFontSize", command: commands.setFontSize(18) },
  { name: "setTextColor", command: commands.setTextColor("#ff0000") },
  {
    name: "setTextBackground",
    command: commands.setTextBackground("#ffff00"),
  },
  {
    name: "setParagraphAlign",
    command: commands.setParagraphAlign("center"),
  },
  {
    name: "setParagraphStyle",
    command: commands.setParagraphStyle("Heading1"),
  },
  {
    name: "setLineSpacing",
    command: commands.setLineSpacing({ rule: "auto", lines: 2 }),
  },
  { name: "setLink", command: commands.setLink("https://example.com") },
  { name: "removeLink", command: commands.removeLink },
  { name: "insertImage", command: commands.insertImage(A_PICTURE) },
  { name: "insertLineBreak", command: commands.insertLineBreak },
  { name: "insertPageBreak", command: commands.insertPageBreak },
  { name: "insertTab", command: commands.insertTab },
  {
    name: "insertTable",
    command: commands.insertTable({ rows: 2, columns: 2 }),
  },
  { name: "lockSelection", command: commands.lockSelection },
  { name: "unlockSelection", command: commands.unlockSelection },
  { name: "undo", command: commands.undo },
  { name: "redo", command: commands.redo },
  { name: "addRowBefore", command: table.addRowBefore },
  { name: "addRowAfter", command: table.addRowAfter },
  { name: "deleteRow", command: table.deleteRow },
  { name: "addColumnBefore", command: table.addColumnBefore },
  { name: "addColumnAfter", command: table.addColumnAfter },
  { name: "deleteColumn", command: table.deleteColumn },
  { name: "deleteTable", command: table.deleteTable },
  { name: "mergeCells", command: table.mergeCells },
  { name: "splitCell", command: table.splitCell },
  { name: "setCellBorders", command: table.setCellBorders("all") },
  {
    name: "setCellBorderColor",
    command: table.setCellBorderColor("#333333"),
  },
  {
    name: "setCellBackground",
    command: table.setCellBackground("#eeeeee"),
  },
  {
    name: "setCellVerticalAlign",
    command: table.setCellVerticalAlign("bottom"),
  },
  {
    name: "setCellPadding",
    command: table.setCellPadding({ left: 6 }),
  },
];

/**
 * Everything else the two entries export, with the reason it answers no such question.
 *
 * `canFormatText`, `canSetLineSpacing`, `canSetParagraphAlign`, `canDecreaseIndent`,
 * `canIncreaseIndent`, `canInsertImage`, `canInsertTable`, `canMergeCells`, `canSplitCell` and
 * `canSetCellBorderColor` are the queries the controls are drawn from; each is checked beside the
 * command it gates in that command's own test file. `canRunCommand` is the question asked of a
 * command the package does not own, so there is no command here for it to disagree with.
 */
const NOT_A_COMMAND: Readonly<Record<string, string>> = {
  IMAGE_FILE_ACCEPT: "the accept string a file picker is given",
  SINGLE_LINE_SPACING: "a line spacing value",
  activeCellBackground: "a query about the selected cells",
  activeCellBorderColor: "a query about the selected cells",
  activeCellPadding: "a query about the selected cells",
  activeCellVerticalAlign: "a query about the selected cells",
  activeFontFamily: "a query about the selection",
  activeFontSize: "a query about the selection",
  activeLineSpacing: "a query about the selection",
  activeLink: "a query about the selection",
  activeLinkSpan: "a query about the selection, the locks left to the commands",
  activeListKind: "a query about the selection",
  activeParagraphAlign: "a query about the selection",
  activeParagraphStyle: "a query about the selection",
  activeTextBackground: "a query about the selection",
  activeTextColor: "a query about the selection",
  canDecreaseIndent: "the query the decrease-indent button is drawn from",
  canAddComment: "the query the add-comment button is drawn from",
  canEditComment:
    "the query the edit and delete buttons of a comment are drawn from",
  canFormatText: "the query the character formatting controls are drawn from",
  canIncreaseIndent: "the query the increase-indent button is drawn from",
  canInsertImage: "the query the image button is drawn from",
  canInsertTable: "the query the insert-table button is drawn from",
  canMergeCells: "the query the merge row is drawn from",
  canRunCommand: "the question asked of a command the package does not own",
  canSetCellBorderColor: "the query the border colour picker is drawn from",
  canSetCellFormatting: "the query the cell layout controls are drawn from",
  canSetLineSpacing: "the query the spacing menu is drawn from",
  canSetLink: "the query the link button is drawn from",
  canSetParagraphAlign: "the query the alignment menu is drawn from",
  canSplitCell: "the query the split row is drawn from",
  documentBodyWidthPx: "a measurement read off the open document",
  documentComments: "the comments displayed alongside the document",
  documentDefaults: "the formatting the document declares",
  documentFontNames: "the fonts the document names",
  documentHasLocked: "a query about the document",
  documentNotes: "the notes displayed after the document",
  documentParagraphStyles: "the styles the document defines",
  editingProtection: "a query about what the editor as a whole may receive",
  fittedExtent: "the rule an oversized image is shrunk by",
  imageFilesIn: "picks the image files out of a picker, clipboard or drag",
  insertImageFiles: "reads the files first, then runs `insertImage`",
  isBoldActive: "a query about the selection",
  isInList: "a query about the selection, answered whatever the lock says",
  isInTable: "the query the table buttons are drawn from",
  isItalicActive: "a query about the selection",
  isStrikeActive: "a query about the selection",
  isUnderlineActive: "a query about the selection",
  readImageFile: "reads one file and gives the size it comes in at",
  selectionLock: "a query about the selection",
  selectionTouchesLocked: "a query about the selection",
};

function sameMarks(
  a: readonly Mark[] | null,
  b: readonly Mark[] | null
): boolean {
  return Mark.sameSet(a ?? [], b ?? []);
}

/** Whether running the command left the state exactly as it stood */
function unchanged(before: EditorState, after: EditorState): boolean {
  return (
    before.doc.eq(after.doc) &&
    before.selection.eq(after.selection) &&
    sameMarks(before.storedMarks, after.storedMarks)
  );
}

/**
 * Runs the command for real and reports both its answer and whether anything moved.
 *
 * Every transaction goes through `apply`, which is where the lock guard stands, so a refused edit
 * shows up here as nothing having changed. Each transaction is applied to the state the one before
 * it produced, so a command that dispatches twice is followed properly.
 */
function attempt(
  state: EditorState,
  command: Command
): { answered: boolean; changed: boolean } {
  let after = state;
  const answered = command(state, (tr) => {
    after = after.apply(tr);
  });
  return { answered, changed: !unchanged(state, after) };
}

describe.each(PROTECTIONS)(
  "every exported command over locked content under %s",
  (protection) => {
    describe.each(PLACES)("with $name", ({ state }) => {
      it.each(CASES)("$name says what dispatching it does", ({ command }) => {
        const before = state(protection);
        const { answered, changed } = attempt(before, command);
        expect(
          changed,
          answered
            ? "the command reported that it applies, and dispatching it changed nothing"
            : "the command reported that it does not apply, and dispatching it changed something"
        ).toBe(answered);
      });
    });
  }
);

describe("comment commands over a real comment anchor", () => {
  it.each([
    [
      "add reply",
      commands.addCommentReply("0", { text: "reply", author: "A" }),
    ],
    ["update", commands.updateComment("0", "updated")],
    ["resolve", commands.setCommentResolved("0", true)],
    ["remove", commands.removeComment("0")],
  ] as const)("refuses to %s inside locked content", (_name, command) => {
    expect(attempt(commentState(true), command)).toEqual({
      answered: false,
      changed: false,
    });
  });

  it.each([
    [
      "add reply",
      commands.addCommentReply("0", { text: "reply", author: "A" }),
    ],
    ["update", commands.updateComment("0", "updated")],
    ["resolve", commands.setCommentResolved("0", true)],
    ["remove", commands.removeComment("0")],
  ] as const)("can %s outside locked content", (_name, command) => {
    expect(attempt(commentState(false), command)).toEqual({
      answered: true,
      changed: true,
    });
  });

  it("selects a locked comment without treating selection as an edit", () => {
    expect(attempt(commentState(true), commands.selectComment("0"))).toEqual({
      answered: true,
      changed: true,
    });
  });

  it.each([
    [
      "add reply",
      commands.addCommentReply("0", { text: "reply", author: "A" }),
    ],
    ["update", commands.updateComment("0", "updated")],
    ["resolve", commands.setCommentResolved("0", true)],
    ["remove", commands.removeComment("0")],
  ] as const)("can %s under the comments protection", (_name, command) => {
    expect(attempt(commentState(false, "comments"), command)).toEqual({
      answered: true,
      changed: true,
    });
  });

  it.each([
    [
      "add reply",
      commands.addCommentReply("0", { text: "reply", author: "A" }),
    ],
    ["update", commands.updateComment("0", "updated")],
    ["resolve", commands.setCommentResolved("0", true)],
    ["remove", commands.removeComment("0")],
  ] as const)("refuses to %s under readOnly", (_name, command) => {
    expect(attempt(commentState(false, "readOnly"), command)).toEqual({
      answered: false,
      changed: false,
    });
  });
});

/**
 * A toggle decides from a query and edits through a filter, and the two have to agree about which
 * pieces are in play. Where they disagreed, the query counted a locked stretch the edit skipped, so
 * the first press turned bold on over the rest and every press after it asked for on again and
 * changed nothing, leaving no way to take it off.
 */
const TOGGLES: readonly CommandCase[] = [
  { name: "toggleBold", command: commands.toggleBold },
  { name: "toggleItalic", command: commands.toggleItalic },
  { name: "toggleStrike", command: commands.toggleStrike },
  { name: "toggleUnderline", command: commands.toggleUnderline },
  { name: "toggleBulletList", command: commands.toggleBulletList },
  { name: "toggleNumberedList", command: commands.toggleNumberedList },
];

/** The state a command leaves behind, every transaction applied through the guard */
function afterRunning(state: EditorState, command: Command): EditorState {
  let after = state;
  command(state, (tr) => {
    after = after.apply(tr);
  });
  return after;
}

describe("every toggle over locked content", () => {
  describe.each(PLACES)("with $name", ({ state }) => {
    it.each(TOGGLES)("$name takes back what it put on", ({ command }) => {
      const before = state("none");
      const once = afterRunning(before, command);
      // A toggle with nothing to reach here says so by changing nothing, which is its own business
      if (before.doc.eq(once.doc)) return;

      const twice = afterRunning(once, command);
      expect(
        twice.doc.eq(before.doc),
        "pressing it twice left the document somewhere else, so what the first press did cannot be undone by the second"
      ).toBe(true);
    });
  });
});

describe("the list of commands above", () => {
  it("covers every export of `./commands` and `./table`", () => {
    const exported = new Set([...Object.keys(commands), ...Object.keys(table)]);
    const covered = new Set([
      ...CASES.map((entry) => entry.name),
      ...Object.keys(NOT_A_COMMAND),
    ]);
    const unanswered = [...exported].filter((name) => !covered.has(name));
    const stale = [...covered].filter((name) => !exported.has(name));

    expect(
      unanswered,
      `exported but not answered for: ${unanswered.join(", ")}\nAdd each to CASES, or to NOT_A_COMMAND with the reason it asks no such question.`
    ).toEqual([]);
    expect(stale, `listed but no longer exported: ${stale.join(", ")}`).toEqual(
      []
    );
  });
});
