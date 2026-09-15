// @vitest-environment jsdom
/**
 * What a lock shuts where the control is a container rather than a mark: a block-level `w:sdt`,
 * and a cell holding one.
 *
 * The cases are built from small bodies so that what each one proves stands in the test itself.
 * `content-controls.docx` is read at the end, since it is the one package carrying all four
 * shapes - a locked control, one locked against deletion alone, a `w:group`, and a control inside
 * a group - as a word processor would write them.
 */
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { AllSelection, type EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx, readFixture } from "../__testing__/docx";
import { rangeOfText, runCommand, select } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import { toggleBold } from "../editor/commands/formatting/editing";
import {
  documentHasLocked,
  selectionLock,
  unlockSelection,
} from "../editor/commands/lockCommands";
import { setParagraphAlign } from "../editor/commands/paragraphCommands";
import {
  createEditorState,
  editorStateForSession,
} from "../editor/createEditor";
import { deleteRow } from "../table";
import { carriesLock, unlockAllowed } from "./locks";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const P = (text: string) => `<w:p>${run(text)}</w:p>`;

/** A control, written the way the fixture writes one: an id, then an optional lock, then its type */
function sdt(
  content: string,
  { id = 1, lock = "", type = "<w:richText/>" } = {}
): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>${locked}${type}</w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

const cell = (inner: string) => `<w:tc><w:p/>${inner}</w:tc>`;

/** An inline control, which stands inside a paragraph rather than around blocks */
function inlineSdt(
  content: string,
  { id = 1, lock = "", type = "" } = {}
): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>${locked}${type}</w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

const table = (...rows: string[]) =>
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
  rows.map((row) => `<w:tr>${row}</w:tr>`).join("") +
  "</w:tbl>";

function opened(body: string): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc);
}

/** Whether the transaction went through: one a guard refused leaves the document as it was */
function applies(state: EditorState, tr: Transaction): boolean {
  return !state.apply(tr).doc.eq(state.doc);
}

/** The position one character into this text, which is where a lock stands between both sides */
function inside(doc: PMNode, needle: string): number {
  return rangeOfText(doc, needle).from + 1;
}

/** Every node of this type holding this text, outermost first, as the stretch each covers */
function spansOf(
  doc: PMNode,
  typeName: string,
  text: string
): { from: number; to: number }[] {
  const spans: { from: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === typeName && node.textContent.includes(text)) {
      spans.push({ from: pos, to: pos + node.nodeSize });
    }
    return true;
  });
  return spans;
}

function spanOf(doc: PMNode, typeName: string, text: string) {
  const first = spansOf(doc, typeName, text)[0];
  if (first === undefined) throw new Error(`no ${typeName} holding ${text}`);
  return first;
}

describe("a block control that shuts its contents", () => {
  const BODY = sdt(P("Shut"), { lock: "sdtContentLocked" }) + P("After");

  it("refuses a character typed inside it", () => {
    const state = opened(BODY);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Shut")))
    ).toBe(false);
  });

  it("refuses a deletion inside it", () => {
    const state = opened(BODY);
    const at = inside(state.doc, "Shut");
    expect(applies(state, state.tr.delete(at, at + 1))).toBe(false);
  });

  it("leaves the blocks around it editable", () => {
    const state = opened(BODY);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "After")))
    ).toBe(true);
  });

  it("refuses a deletion that crosses its edge", () => {
    const state = opened(
      P("Before") + sdt(P("Shut"), { lock: "contentLocked" })
    );
    expect(
      applies(
        state,
        state.tr.delete(
          inside(state.doc, "Before"),
          inside(state.doc, "Shut") + 1
        )
      )
    ).toBe(false);
  });

  /**
   * The two clauses are attributes of the control, and a step may rewrite a node's attributes
   * where it stands, so the lock has to be shut against those very steps as a cell's is: lifting
   * a lock stays the business of the command that carries the pass for it.
   */
  it("refuses a step that puts one of its clauses down by itself", () => {
    const state = opened(BODY);
    const at = spanOf(state.doc, "sdtBlock", "Shut").from;
    const node = state.doc.nodeAt(at);
    if (!node) throw new Error("no control");

    expect(
      applies(state, state.tr.setNodeAttribute(at, "contentsLocked", false))
    ).toBe(false);
    expect(
      applies(state, state.tr.setNodeAttribute(at, "deletionLocked", false))
    ).toBe(false);
    expect(
      applies(
        state,
        state.tr.setNodeMarkup(at, null, {
          ...node.attrs,
          contentsLocked: false,
          deletionLocked: false,
        })
      )
    ).toBe(false);
  });

  it("takes the same step once it carries the unlocking pass", () => {
    const state = opened(BODY);
    const at = spanOf(state.doc, "sdtBlock", "Shut").from;
    const tr = state.tr
      .setMeta(unlockAllowed, true)
      .setNodeAttribute(at, "contentsLocked", false);
    expect(applies(state, tr)).toBe(true);
  });

  it("carries a lock the document can be asked about", () => {
    const state = opened(BODY);
    expect(carriesLock(state.doc.child(0))).toBe(true);
    expect(documentHasLocked(state.doc)).toBe(true);
  });
});

describe("taking a block control away whole", () => {
  const bodyWith = (lock: string) => sdt(P("Whole"), { lock }) + P("After");

  it.each([
    ["nothing at all", "", true],
    ["contentLocked, which shuts the contents alone", "contentLocked", true],
    ["sdtLocked, which shuts the wrapper alone", "sdtLocked", false],
    ["sdtContentLocked, which shuts both", "sdtContentLocked", false],
  ])("is %s -> %s", (_name, lock, allowed) => {
    const state = opened(bodyWith(lock));
    const span = spanOf(state.doc, "sdtBlock", "Whole");
    expect(applies(state, state.tr.delete(span.from, span.to))).toBe(allowed);
  });
});

describe("a control standing inside another carrier", () => {
  it("shuts an inline control inside an open block control", () => {
    const state = opened(
      sdt(
        `<w:p>${run("open ")}` +
          `<w:sdt><w:sdtPr><w:id w:val="9"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>` +
          `<w:sdtContent>${run("shut")}</w:sdtContent></w:sdt></w:p>`
      )
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "shut")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "open ")))
    ).toBe(true);
  });

  it("shuts a locked block control standing inside a cell", () => {
    const state = opened(
      "<w:tbl>" +
        '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
        "<w:tr>" +
        cell(sdt(P("InCell"), { lock: "sdtContentLocked" })) +
        `<w:tc>${P("Beside")}</w:tc>` +
        "</w:tr></w:tbl>"
    );
    const table = spanOf(state.doc, "table", "InCell");

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "InCell")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Beside")))
    ).toBe(true);
    // Taking the table away covers the control whole, which its deletion clause refuses
    expect(applies(state, state.tr.delete(table.from, table.to))).toBe(false);
  });
});

describe("a block control holding a table", () => {
  const BODY = sdt(table(cell(P("InTable")), cell(P("Second"))), {
    lock: "sdtContentLocked",
  });

  it("refuses a character typed in one of its cells", () => {
    const state = opened(BODY);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "InTable")))
    ).toBe(false);
  });

  it("refuses a paragraph command and a mark command inside it", () => {
    const state = opened(BODY);
    const at = inside(state.doc, "InTable");
    const caret = select(state, at);
    expect(setParagraphAlign("center")(caret)).toBe(false);
    expect(toggleBold(select(state, at, at + 3))).toBe(false);
  });

  it("refuses a row deletion inside it", () => {
    const state = opened(BODY);
    expect(deleteRow(select(state, inside(state.doc, "InTable")))).toBe(false);
  });
});

describe("a control standing inside a lock", () => {
  it("shuts an open block control a locked cell holds", () => {
    const state = opened(
      table(sdt(cell(sdt(P("Inner"), { id: 2 })), { lock: "sdtContentLocked" }))
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Inner")))
    ).toBe(false);
  });

  it("shuts an open inline control a locked inline control holds", () => {
    const state = opened(
      `<w:p>${inlineSdt(inlineSdt(run("inner"), { id: 2 }), {
        id: 1,
        lock: "sdtContentLocked",
      })}</w:p>`
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "inner")))
    ).toBe(false);
  });
});

describe("a w:group control", () => {
  it("shuts the blocks it holds although it states no lock", () => {
    const state = opened(
      sdt(P("Grouped"), { type: "<w:group/>" }) + P("After")
    );

    // The group is carried apart from the lock, which is what leaves it nothing to lift
    expect(state.doc.child(0).attrs.group).toBe(true);
    expect(state.doc.child(0).attrs.contentsLocked).toBe(false);
    expect(state.doc.child(0).attrs.deletionLocked).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Grouped")))
    ).toBe(false);
  });

  it("may still be taken away whole, since a group states nothing about that", () => {
    const state = opened(
      sdt(P("Grouped"), { type: "<w:group/>" }) + P("After")
    );
    const span = spanOf(state.doc, "sdtBlock", "Grouped");
    expect(applies(state, state.tr.delete(span.from, span.to))).toBe(true);
  });

  it("shuts the runs it holds where it stands inside a paragraph", () => {
    const state = opened(
      `<w:p>${run("ahead")}` +
        '<w:sdt><w:sdtPr><w:id w:val="5"/><w:group/></w:sdtPr>' +
        `<w:sdtContent>${run("grouped")}</w:sdtContent></w:sdt>${run("d")}</w:p>`
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "grouped")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "ahead")))
    ).toBe(true);
  });

  it("is superseded by a block control standing inside it", () => {
    const state = opened(
      sdt(P("Direct") + sdt(P("Inner"), { id: 2 }), {
        id: 1,
        type: "<w:group/>",
      })
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Direct")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Inner")))
    ).toBe(true);
  });

  it("is superseded by an inline control standing inside it", () => {
    const state = opened(
      sdt(`<w:p>${run("plain ")}${inlineSdt(run("inner"), { id: 2 })}</w:p>`, {
        id: 1,
        type: "<w:group/>",
      })
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "inner")))
    ).toBe(true);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "plain ")))
    ).toBe(false);
  });

  it("shuts the cell it wraps", () => {
    const state = opened(
      table(sdt(cell(P("Grouped")), { type: "<w:group/>" }), cell(P("Beside")))
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Grouped")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "Beside")))
    ).toBe(true);
  });

  /**
   * A lock states no exception of its own (§17.18.49), so an open control inside a group inside a
   * lock opens the group and is still shut by the lock standing around both.
   */
  it("opens nothing where a lock stands around it", () => {
    const state = opened(
      sdt(
        sdt(
          `<w:p>${run("plain ")}${inlineSdt(run("inner"), { id: 3 })}</w:p>`,
          { id: 2, type: "<w:group/>" }
        ),
        { id: 1, lock: "sdtContentLocked" }
      )
    );

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "inner")))
    ).toBe(false);
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, "plain ")))
    ).toBe(false);
  });
});

describe("lifting the lock off a block control", () => {
  function unlockedAll(body: string) {
    const opened = importDocx(makeDocx(body));
    const state = editorStateForSession(opened);
    const all = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
    return { after: runCommand(all, unlockSelection), session: opened.session };
  }

  it("opens both clauses and writes the control back without its w:lock", () => {
    const { after, session } = unlockedAll(
      sdt(P("Shut"), { id: 42, lock: "sdtContentLocked" })
    );
    const control = after.doc.child(0);

    expect(control.attrs.contentsLocked).toBe(false);
    expect(control.attrs.deletionLocked).toBe(false);
    expect(
      applies(after, after.tr.insertText("x", inside(after.doc, "Shut")))
    ).toBe(true);

    const xml = documentXmlOf(after.doc, session);
    expect(xml).not.toContain("<w:lock");
    // The control itself stays, under the id the file gave it (§17.5.2.18)
    expect(xml).toContain('<w:id w:val="42"/>');
  });

  it("offers nothing to lift on a group, whose contents no w:lock shut", () => {
    const state = editorStateForSession(
      importDocx(makeDocx(sdt(P("Grouped"), { id: 43, type: "<w:group/>" })))
    );
    const all = state.apply(state.tr.setSelection(new AllSelection(state.doc)));

    expect(unlockSelection(all, () => undefined)).toBe(false);
    expect(all.doc.child(0).attrs.group).toBe(true);
    // Not `none`: the selection stands in contents that refuse every edit
    expect(selectionLock(all)).toBe("shut");
  });
});

describe("the locks content-controls.docx carries", () => {
  const SHUT = "The contents of this control may not be edited";
  const KEPT = "The contents of this control may be retyped";
  const IN_GROUP = "The inner control holds the only paragraph of the pair.";

  function fixtureState(): EditorState {
    return editorStateForSession(
      importDocx(readFixture("content-controls.docx"))
    );
  }

  it("refuses an edit inside the control that shuts its contents", () => {
    const state = fixtureState();
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, SHUT)))
    ).toBe(false);
  });

  it("lets the control locked against deletion alone be retyped, not removed", () => {
    const state = fixtureState();
    const span = spanOf(state.doc, "sdtBlock", KEPT);

    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, KEPT)))
    ).toBe(true);
    expect(applies(state, state.tr.delete(span.from, span.to))).toBe(false);
  });

  it("lets the control inside the group be edited, and says the document is locked", () => {
    const state = fixtureState();
    expect(
      applies(state, state.tr.insertText("x", inside(state.doc, IN_GROUP)))
    ).toBe(true);
    expect(documentHasLocked(state.doc)).toBe(true);
  });
});
