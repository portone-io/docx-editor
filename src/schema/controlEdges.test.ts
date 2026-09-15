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
