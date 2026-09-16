// @vitest-environment jsdom
/**
 * The edge of a block-level content control, put to the test through the keys that meet it.
 *
 * Each case runs the command the key is bound to and applies whatever it dispatches through the
 * state itself, which is where the guard sits (`editor/plugins/lockedContent`), so what the tests
 * read is what a user gets rather than what a command intended.
 */
import { baseKeymap, chainCommands } from "prosemirror-commands";
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  NodeSelection,
  TextSelection,
} from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx } from "../__testing__/docx";
import { select } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import type { SessionStore } from "../docx/session";
import { editorStateForSession } from "../editor/createEditor";
import { docxKeymap } from "../editor/plugins/keymap";
import { docxSchema } from "./index";

const P = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function sdt(content: string, { id = 1, lock = "" } = {}): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>${locked}<w:richText/></w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

function stateOf(body: string): EditorState {
  return editorStateForSession(importDocx(makeDocx(body)));
}

function openedWith(body: string): {
  state: EditorState;
  session: SessionStore;
} {
  const opened = importDocx(makeDocx(body));
  return { state: editorStateForSession(opened), session: opened.session };
}

/** Every top-level block, a control spelled out as the blocks it holds */
function shape(doc: PMNode): string {
  const blocks: string[] = [];
  doc.forEach((block) => {
    if (block.type.name !== "sdtBlock") {
      blocks.push(`${block.type.name}(${block.textContent})`);
      return;
    }
    const inside: string[] = [];
    block.forEach((child) => {
      inside.push(`${child.type.name}(${child.textContent})`);
    });
    blocks.push(`control[${inside.join(",")}]`);
  });
  return blocks.join(" ");
}

/** The position just inside the start of the block whose whole text reads this */
function startOf(doc: PMNode, text: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isTextblock && node.textContent === text) {
      found = pos + 1;
    }
  });
  if (found < 0) throw new Error(`no block reading ${text}`);
  return found;
}

function endOf(doc: PMNode, text: string): number {
  return startOf(doc, text) + text.length;
}

const backspace = chainCommands(
  docxKeymap.Backspace as Command,
  baseKeymap.Backspace as Command
);
const del = chainCommands(
  docxKeymap.Delete as Command,
  baseKeymap.Delete as Command
);

/** The state the key leaves behind, every transaction applied through the guard */
function press(state: EditorState, key: Command): EditorState {
  let after = state;
  key(state, (tr) => {
    after = after.apply(tr);
  });
  return after;
}

describe("a join across a control's edge", () => {
  it("does nothing on Backspace at the start of the first block", () => {
    const state = stateOf(P("Outside") + sdt(P("First") + P("Second")));
    const before = select(state, startOf(state.doc, "First"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First),paragraph(Second)]"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("does nothing when the control holds a single paragraph of text", () => {
    const state = stateOf(P("Outside") + sdt(P("First")));
    const before = select(state, startOf(state.doc, "First"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First)]"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("does nothing when the control is the first block of the document", () => {
    const state = stateOf(sdt(P("First")) + P("Outside"));
    const before = select(state, startOf(state.doc, "First"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "control[paragraph(First)] paragraph(Outside)"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("does nothing on Delete at the end of the last block", () => {
    const state = stateOf(sdt(P("First") + P("Second")) + P("Outside"));
    const before = select(state, endOf(state.doc, "Second"));

    const after = press(before, del);

    expect(shape(after.doc)).toBe(
      "control[paragraph(First),paragraph(Second)] paragraph(Outside)"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("refuses a selection that runs from outside a control into it", () => {
    const state = stateOf(P("Outside") + sdt(P("First") + P("Second")));
    const before = select(
      state,
      startOf(state.doc, "Outside") + 4,
      startOf(state.doc, "First") + 2
    );

    const after = before.apply(before.tr.deleteSelection());

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First),paragraph(Second)]"
    );
  });

  it("refuses typing over such a selection, which would carry the text out", () => {
    const state = stateOf(P("Outside") + sdt(P("First") + P("Second")));
    const from = startOf(state.doc, "Outside") + 4;
    const to = startOf(state.doc, "First") + 2;
    const before = select(state, from, to);

    const after = before.apply(before.tr.insertText("X", from, to));

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First),paragraph(Second)]"
    );
  });

  it("refuses a selection that runs from inside a control out of it", () => {
    const state = stateOf(sdt(P("First") + P("Second")) + P("Outside"));
    const before = select(
      state,
      startOf(state.doc, "Second") + 3,
      startOf(state.doc, "Outside") + 4
    );

    const after = before.apply(before.tr.deleteSelection());

    expect(shape(after.doc)).toBe(
      "control[paragraph(First),paragraph(Second)] paragraph(Outside)"
    );
  });
});

describe("what the edge leaves open", () => {
  it("lets a selection covering the control whole take it away with its contents", () => {
    const state = stateOf(P("Outside") + sdt(P("First")) + P("Tail"));
    const before = select(
      state,
      startOf(state.doc, "Outside") + 4,
      startOf(state.doc, "Tail") + 2
    );

    const after = before.apply(before.tr.deleteSelection());

    expect(shape(after.doc)).toBe("paragraph(Outsil)");
  });

  it("lets the control be selected as a node and deleted", () => {
    const state = stateOf(P("Outside") + sdt(P("First")) + P("Tail"));
    const at = state.doc.child(0).nodeSize;
    const before = state.apply(
      state.tr.setSelection(NodeSelection.create(state.doc, at))
    );

    const after = before.apply(before.tr.deleteSelection());

    expect(shape(after.doc)).toBe("paragraph(Outside) paragraph(Tail)");
  });

  it("keeps a paragraph split inside the control, at either end of it", () => {
    const state = stateOf(sdt(P("First")) + P("Outside"));
    const atEnd = select(state, endOf(state.doc, "First"));

    const split = press(atEnd, docxKeymap.Enter as Command);

    expect(shape(split.doc)).toBe(
      "control[paragraph(First),paragraph()] paragraph(Outside)"
    );

    const atStart = select(state, startOf(state.doc, "First"));
    const opened = press(atStart, docxKeymap.Enter as Command);

    expect(shape(opened.doc)).toBe(
      "control[paragraph(),paragraph(First)] paragraph(Outside)"
    );
  });

  it("lets an edit inside the control go through untouched", () => {
    const state = stateOf(P("Outside") + sdt(P("First")));
    const at = endOf(state.doc, "First");
    const before = select(state, at);

    const after = before.apply(before.tr.insertText("!", at));

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First!)]"
    );
  });
});

describe("a control holding one empty paragraph", () => {
  it("goes whole on Backspace, leaving the caret where it stood", () => {
    const state = stateOf(P("Outside") + sdt("<w:p/>") + P("Tail"));
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe("paragraph(Outside) paragraph(Tail)");
    expect(after.selection.from).toBe(endOf(after.doc, "Outside"));
  });

  it("goes whole on Delete as well", () => {
    const state = stateOf(P("Outside") + sdt("<w:p/>") + P("Tail"));
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, del);

    expect(shape(after.doc)).toBe("paragraph(Outside) paragraph(Tail)");
  });

  it("takes the control around it too, since block+ admits nothing empty", () => {
    const state = stateOf(P("Outside") + sdt(sdt("<w:p/>", { id: 2 })));
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe("paragraph(Outside)");
  });

  it("leaves a one-cell table inside the control standing, cell and all", () => {
    const cell = "<w:tc><w:tcPr/><w:p/></w:tc>";
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr>${cell}</w:tr></w:tbl>`;
    const state = stateOf(P("Outside") + sdt(table) + P("Tail"));
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[table()] paragraph(Tail)"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("is refused where the control says it may not be deleted", () => {
    const state = stateOf(
      P("Outside") + sdt("<w:p/>", { lock: "sdtLocked" }) + P("Tail")
    );
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph()] paragraph(Tail)"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("goes whole where the control locks its contents alone, which the deletion clause allows", () => {
    const state = stateOf(
      P("Outside") + sdt("<w:p/>", { lock: "contentLocked" }) + P("Tail")
    );
    const before = select(state, startOf(state.doc, ""));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe("paragraph(Outside) paragraph(Tail)");
  });

  it("leaves a control holding a paragraph with text alone", () => {
    const state = stateOf(P("Outside") + sdt(P("First")));
    const before = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, startOf(state.doc, "First"))
      )
    );

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First)]"
    );
  });
});

/**
 * A control holding nothing is an atom rather than a container (`docx/importSdtBlock`), so no edge
 * of it stands between two blocks: `controlsAround` walks the ancestors of a position and an atom
 * is never one of them. A keystroke beside it passes the caret over it rather than joining
 * anything, and only a selection covering it takes it away, which is the deletion clause's
 * question (`./locks`).
 */
describe("a control holding nothing", () => {
  const empty = (lock = "", id = 1) => sdt("", { id, lock });
  /** A paragraph drawn as a blank line, which is not the control and goes as any blank line does */
  const BLANK = "<w:p/>";

  it("passes the caret over it on Backspace from the block after it", () => {
    const state = stateOf(P("Head") + empty() + P("Tail"));
    const before = select(state, startOf(state.doc, "Tail"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe("paragraph(Head) sdtEmpty() paragraph(Tail)");
    expect(after.selection.from).toBe(endOf(state.doc, "Head"));
  });

  it("passes the caret over it on Delete from the block before it", () => {
    const state = stateOf(P("Head") + empty() + P("Tail"));
    const before = select(state, endOf(state.doc, "Head"));

    const after = press(before, del);

    expect(shape(after.doc)).toBe("paragraph(Head) sdtEmpty() paragraph(Tail)");
    expect(after.selection.from).toBe(startOf(state.doc, "Tail"));
  });

  it("passes a run of them in one keystroke", () => {
    const state = stateOf(P("Head") + empty() + empty("", 2) + P("Tail"));
    const before = select(state, startOf(state.doc, "Tail"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Head) sdtEmpty() sdtEmpty() paragraph(Tail)"
    );
    expect(after.selection.from).toBe(endOf(state.doc, "Head"));
  });

  it("does nothing on Backspace where it stands at the start of the document", () => {
    const state = stateOf(empty() + P("Tail"));
    const before = select(state, startOf(state.doc, "Tail"));

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(shape(state.doc));
    expect(after.selection.from).toBe(before.selection.from);
  });

  /**
   * The blank line is not the control, so it goes the way a blank line goes anywhere else, and the
   * caret lands past the control rather than on the line that went.
   */
  it("takes the caret's own blank line away on Backspace", () => {
    const state = stateOf(P("Head") + empty() + BLANK + P("Tail"));
    const before = select(state, startOf(state.doc, "Tail") - 2);

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe("paragraph(Head) sdtEmpty() paragraph(Tail)");
    expect(after.selection.from).toBe(endOf(state.doc, "Head"));
  });

  it("takes the caret's own blank line away on Delete", () => {
    const state = stateOf(P("Head") + BLANK + empty() + P("Tail"));
    const before = select(state, endOf(state.doc, "Head") + 2);

    const after = press(before, del);

    expect(shape(after.doc)).toBe("paragraph(Head) sdtEmpty() paragraph(Tail)");
    expect(after.selection.from).toBe(startOf(after.doc, "Tail"));
  });

  it("keeps the blank line where nothing stands beyond the control", () => {
    const state = stateOf(empty() + BLANK + P("Tail"));
    const before = select(state, startOf(state.doc, "Tail") - 2);

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(shape(state.doc));
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("does nothing on Delete where it stands at the end of the document", () => {
    const state = stateOf(P("Head") + empty());
    const before = select(state, endOf(state.doc, "Head"));

    const after = press(before, del);

    expect(shape(after.doc)).toBe(shape(state.doc));
    expect(after.selection.from).toBe(before.selection.from);
  });

  it.each([
    ["nothing at all", "", true],
    ["contentLocked, which shuts its contents alone", "contentLocked", true],
    ["sdtLocked, which shuts it against deletion", "sdtLocked", false],
  ])(
    "goes with a stretch that covers it where it states %s -> %s",
    (_n, lock, goes) => {
      const state = stateOf(P("Head") + empty(lock) + P("Tail"));
      const before = select(
        state,
        endOf(state.doc, "Head"),
        startOf(state.doc, "Tail")
      );

      const after = before.apply(before.tr.deleteSelection());

      expect(shape(after.doc)).toBe(
        goes
          ? "paragraph(HeadTail)"
          : "paragraph(Head) sdtEmpty() paragraph(Tail)"
      );
    }
  );

  /**
   * The node holds no spot a caret could stand in, and the editor registers no gap cursor
   * (`editor/createEditor`), so what stands either side of it is where a caret goes. It is not
   * selectable either: it draws nothing, so a node selection would leave the user in a state with
   * nothing on screen to show it, and the next keystroke would replace a control they cannot see.
   */
  it("leaves the caret free where two of them stand next to each other", () => {
    const state = stateOf(empty() + empty("", 2) + P("Tail"));

    expect(shape(state.doc)).toBe("sdtEmpty() sdtEmpty() paragraph(Tail)");
    expect(NodeSelection.isSelectable(state.doc.child(0))).toBe(false);
    expect(state.selection.empty).toBe(true);
    expect(state.selection.$from.parent.textContent).toBe("Tail");

    const caret = select(state, startOf(state.doc, "Tail"));
    expect(caret.selection.empty).toBe(true);
    expect(caret.selection.$from.parent.textContent).toBe("Tail");
  });
});

/**
 * A control a paragraph holds with nothing inside it is an inline atom (`docx/wrappers`),
 * so it stands between the runs rather than around them. A keystroke beside it passes over it and
 * takes what stands beyond it, as if it were not there, and only a selection covering it takes the
 * control away, which is the deletion clause's question (`./locks`).
 */
describe("a control a paragraph holds with nothing inside it", () => {
  const empty = (lock = "", id = 1) =>
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>` +
    (lock === "" ? "" : `<w:lock w:val="${lock}"/>`) +
    "</w:sdtPr><w:sdtContent></w:sdtContent></w:sdt>";
  const run = (text: string) =>
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const line = (inline: string) => `<w:p>${inline}</w:p>`;

  /** The inline nodes of the one paragraph, a control spelled out as the nothing it holds */
  function inlines(doc: PMNode): string {
    const parts: string[] = [];
    doc.child(0).forEach((child) => {
      parts.push(child.isText ? `text(${child.text})` : `${child.type.name}()`);
    });
    return parts.join(" ");
  }

  /** Where the one control stands, which holds no text to be found by */
  function controlAt(doc: PMNode): number {
    let found = -1;
    doc.descendants((node, pos) => {
      if (found < 0 && node.type.name === "sdtEmptyInline") found = pos;
    });
    if (found < 0) throw new Error("no control holding nothing");
    return found;
  }

  it("takes the character before it on Backspace from the text after it", () => {
    const state = stateOf(line(run("ab") + empty() + run("cd")));
    const before = select(state, controlAt(state.doc) + 1);

    const after = press(before, backspace);

    expect(inlines(after.doc)).toBe("text(a) sdtEmptyInline() text(cd)");
    expect(after.selection.from).toBe(controlAt(after.doc) + 1);
  });

  it("takes the character after it on Delete from the text before it", () => {
    const state = stateOf(line(run("ab") + empty() + run("cd")));
    const before = select(state, controlAt(state.doc));

    const after = press(before, del);

    expect(inlines(after.doc)).toBe("text(ab) sdtEmptyInline() text(d)");
    expect(after.selection.from).toBe(controlAt(after.doc));
  });

  it("passes a run of them in one keystroke", () => {
    const state = stateOf(line(run("ab") + empty() + empty("", 2) + run("cd")));
    const before = select(state, controlAt(state.doc) + 2);

    const after = press(before, backspace);

    expect(inlines(after.doc)).toBe(
      "text(a) sdtEmptyInline() sdtEmptyInline() text(cd)"
    );
    expect(after.selection.from).toBe(controlAt(after.doc) + 2);
  });

  /**
   * The key takes what stands beyond the run, and at the start of a line that is the line break
   * itself: the paragraph joins the one above it as it would with no control standing there.
   */
  it("joins the line into the one above where it opens one", () => {
    const state = stateOf(P("Above") + line(empty() + run("cd")));
    const before = select(state, controlAt(state.doc) + 1);

    const after = press(before, backspace);

    expect(after.doc.childCount).toBe(1);
    expect(inlines(after.doc)).toBe("text(Above) sdtEmptyInline() text(cd)");
  });

  /**
   * What the key takes is the character beyond the run, never the control: the lock has nothing to
   * refuse here, and the control the file keeps stays where it stood.
   */
  it("leaves a control locked against deletion standing", () => {
    const state = stateOf(line(run("ab") + empty("sdtLocked") + run("cd")));
    const before = select(state, controlAt(state.doc) + 1);

    const after = press(before, backspace);

    expect(inlines(after.doc)).toBe("text(a) sdtEmptyInline() text(cd)");
  });

  /**
   * The node holds nothing, so there is no inside for a character to land in: what a keystroke
   * beside it writes stands beside it, wearing none of what the control states.
   */
  it("types beside it rather than inside it", () => {
    const state = stateOf(line(run("ab") + empty() + run("cd")));
    const at = controlAt(state.doc);
    const before = select(state, at + 1);

    const after = before.apply(before.tr.insertText("X", at + 1));
    const typed = after.doc.child(0).child(2);

    expect(after.doc.child(0).child(1).type.name).toBe("sdtEmptyInline");
    expect(typed.text).toBe("X");
    expect(typed.marks.map((mark) => mark.type.name)).toEqual([]);
  });

  it.each([
    ["nothing at all", "", true],
    ["contentLocked, which shuts its contents alone", "contentLocked", true],
    ["sdtLocked, which shuts it against deletion", "sdtLocked", false],
  ])(
    "goes with a stretch that covers it where it states %s -> %s",
    (_n, lock, goes) => {
      const state = stateOf(line(run("ab") + empty(lock) + run("cd")));
      const at = controlAt(state.doc);
      const before = select(state, at, at + 1);

      const after = before.apply(before.tr.deleteSelection());

      expect(inlines(after.doc)).toBe(
        goes ? "text(abcd)" : "text(ab) sdtEmptyInline() text(cd)"
      );
    }
  );

  /**
   * It takes no width, so a node selection of it would show nothing at all, and the keystroke
   * after it would replace a control the user cannot see. Only a stretch covering it reaches it,
   * which is the deletion clause's question (`./locks`).
   */
  it("is no node a selection can settle on", () => {
    const state = stateOf(line(run("ab") + empty() + run("cd")));
    const control = state.doc.child(0).child(1);

    expect(control.type.name).toBe("sdtEmptyInline");
    expect(NodeSelection.isSelectable(control)).toBe(false);
  });
});

describe("a control on the clipboard", () => {
  it("lands inside the control when the caret stands in one", () => {
    const state = stateOf(P("Outside") + sdt(P("First")));
    const at = endOf(state.doc, "First");
    const before = select(state, at);
    const pasted = docxSchema.nodes.paragraph.create(
      null,
      docxSchema.text("Pasted")
    );

    const after = before.apply(
      before.tr.replaceSelection(new Slice(Fragment.from(pasted), 0, 0))
    );

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First),paragraph(Pasted)]"
    );
  });

  it("holds the edge between a control and the copy of it standing next to it", () => {
    const state = stateOf(P("Outside") + sdt(P("First")));
    const control = state.doc.child(1);
    const at = select(state, endOf(state.doc, "Outside"));
    const pasted = at.apply(
      at.tr.replaceSelection(new Slice(Fragment.from(control), 0, 0))
    );
    // Copying a control copies its `key`, so the two are told apart by where they stand
    expect(pasted.doc.child(1).attrs.key).toBe(pasted.doc.child(2).attrs.key);
    const second = pasted.doc.child(0).nodeSize + pasted.doc.child(1).nodeSize;
    const before = select(pasted, second + 2);

    const after = press(before, backspace);

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First)] control[paragraph(First)]"
    );
    expect(after.selection.from).toBe(before.selection.from);
  });

  it("keeps the wrapper of a whole control, under a w:id of its own", () => {
    const opened = openedWith(P("Outside") + sdt(P("First")));
    const { state } = opened;
    const control = state.doc.child(1);
    const at = endOf(state.doc, "Outside");
    const before = select(state, at);

    const after = before.apply(
      before.tr.replaceSelection(new Slice(Fragment.from(control), 0, 0))
    );

    expect(shape(after.doc)).toBe(
      "paragraph(Outside) control[paragraph(First)] control[paragraph(First)]"
    );
    const written = documentXmlOf(after.doc, opened.session);
    const ids = [...written.matchAll(/<w:id w:val="(\d+)"\/>/g)].map(
      (match) => match[1]
    );
    expect(ids).toHaveLength(2);
    expect(ids[0]).not.toBe(ids[1]);
  });
});
