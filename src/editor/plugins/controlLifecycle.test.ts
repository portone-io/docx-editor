// @vitest-environment jsdom
/**
 * The two properties a control carries about outliving an edit of its contents.
 *
 * Each case edits through a state built the way the editor builds one, so the appended transaction
 * is judged by the guards exactly as it is on screen.
 */

import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx, makeNotesDocx } from "../../__testing__/docx";
import { posOfText, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { storyOf } from "../../docx/story";
import { W_NS } from "../../ooxml/xml";
import { storyKey } from "../../schema/stories";
import { undo } from "../commands/historyCommands";
import { editorStateForSession } from "../createEditor";
import { documentOf, storyDocument } from "../editorDocument";
import { storyEditorState } from "../stories/storyState";
import { NO_CAPABILITY } from "../stories/storyView";

const P = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const R = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

function sdt(content: string, props: string, id = 1): string {
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>${props}<w:richText/></w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

const TEMPORARY = "<w:temporary/>";
const PLACEHOLDER = "<w:showingPlcHdr/>";

interface Opened {
  state: EditorState;
  session: ReturnType<typeof importDocx>["session"];
}

function open(body: string): Opened {
  const opened = importDocx(makeDocx(body));
  return { state: editorStateForSession(opened), session: opened.session };
}

/** Types one character just inside the first text node reading this, and applies what follows */
function type(opened: Opened, needle: string, text: string): EditorState {
  const at = posOfText(opened.state.doc, needle);
  const placed = select(opened.state, at);
  return placed.apply(placed.tr.insertText(text, at));
}

function blockTypes(doc: PMNode): string[] {
  const names: string[] = [];
  doc.forEach((block) => {
    names.push(block.type.name);
  });
  return names;
}

function exported(state: EditorState, opened: Opened): string {
  return documentXmlOf(state.doc, opened.session);
}

describe("a control that says it goes once its contents are edited", () => {
  it("loses its wrapper on the first edit inside it, keeping the blocks it held", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), TEMPORARY));

    const after = type(opened, "Inside", "!");

    expect(blockTypes(after.doc)).toEqual(["paragraph", "paragraph"]);
    expect(after.doc.textContent).toBe("OutsideI!nside");
    expect(exported(after, opened)).not.toContain("<w:sdt>");
  });

  it("stands where the edit was somewhere else", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), TEMPORARY));

    const after = type(opened, "Outside", "!");

    expect(blockTypes(after.doc)).toEqual(["paragraph", "sdtBlock"]);
    expect(exported(after, opened)).toContain("<w:temporary/>");
  });

  it("takes the wrapper and the edit back together on one undo", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), TEMPORARY));
    const after = type(opened, "Inside", "!");

    let back = after;
    undo(after, (tr) => {
      back = back.apply(tr);
    });

    expect(blockTypes(back.doc)).toEqual(["paragraph", "sdtBlock"]);
    expect(back.doc.textContent).toBe("OutsideInside");
  });

  it("loses the mark where the control is an inline one", () => {
    const opened = open(
      `<w:p>${R("Plain")}${sdt(R("Inside"), TEMPORARY)}</w:p>`
    );

    const after = type(opened, "Inside", "!");

    expect(exported(after, opened)).not.toContain("<w:sdt>");
    expect(after.doc.textContent).toBe("PlainI!nside");
  });

  it("stays where its lock says it may not be deleted", () => {
    const opened = open(
      P("Outside") + sdt(P("Inside"), `${TEMPORARY}<w:lock w:val="sdtLocked"/>`)
    );

    const after = type(opened, "Inside", "!");

    expect(blockTypes(after.doc)).toEqual(["paragraph", "sdtBlock"]);
  });
});

describe("a control whose contents are placeholder text", () => {
  it("drops the flag from its opening XML on the first edit inside it", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), PLACEHOLDER));

    const after = type(opened, "Inside", "!");

    const control = after.doc.child(1);
    expect(control.type.name).toBe("sdtBlock");
    expect(control.attrs.sdtPrefix).not.toContain("showingPlcHdr");
    expect(control.attrs.showingPlaceholder).toBe(false);
    expect(control.textContent).toBe("I!nside");
  });

  it("keeps the placeholder text itself, which is edited like any other text", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), PLACEHOLDER));

    const after = type(opened, "Inside", "!");

    const written = exported(after, opened);
    expect(written).toContain("<w:sdt>");
    expect(written).not.toContain("showingPlcHdr");
    expect(written).toContain("I!nside");
  });

  it("leaves the flag standing until something inside is edited", () => {
    const opened = open(P("Outside") + sdt(P("Inside"), PLACEHOLDER));

    const after = type(opened, "Outside", "!");

    expect(exported(after, opened)).toContain("<w:showingPlcHdr/>");
  });

  it("drops the flag from an inline control too", () => {
    const opened = open(
      `<w:p>${R("Plain")}${sdt(R("Inside"), PLACEHOLDER)}</w:p>`
    );

    const after = type(opened, "Inside", "!");

    const written = exported(after, opened);
    expect(written).toContain("<w:sdt>");
    expect(written).not.toContain("showingPlcHdr");
  });
});

/**
 * Both properties are about an edit of the control's contents, and a control holding nothing has
 * none: nothing can be typed inside it, so neither ever acts and both ride back out in the prefix.
 */
describe("a control holding nothing", () => {
  it.each([
    ["w:temporary", TEMPORARY, "<w:temporary/>"],
    ["w:showingPlcHdr", PLACEHOLDER, "<w:showingPlcHdr/>"],
  ])("keeps %s while the document is edited around it", (_name, props, xml) => {
    const opened = open(P("Outside") + sdt("", props));

    const after = type(opened, "Outside", "!");

    expect(blockTypes(after.doc)).toEqual(["paragraph", "sdtEmpty"]);
    expect(exported(after, opened)).toContain(xml);
  });

  /**
   * A step that rewrites the node where it stands covers the very stretch the walk asks about,
   * since the node holds nothing and its two ends are one and the same spot. Settling it there
   * would take away the opening XML it goes back out as, which is a file that cannot be written.
   */
  it("keeps w:temporary where a step rewrites the node where it stands", () => {
    const opened = open(P("Outside") + sdt("", TEMPORARY));
    const at = opened.state.doc.child(0).nodeSize;
    const control = opened.state.doc.child(1);

    const after = opened.state.apply(
      opened.state.tr.setNodeMarkup(at, undefined, {
        ...control.attrs,
        key: 99,
      })
    );

    expect(after.doc.child(1).attrs.sdtPrefix).toBe(control.attrs.sdtPrefix);
    expect(exported(after, opened)).toContain("<w:temporary/>");
  });
});

describe("a control a lock says may not be deleted", () => {
  it("keeps its wrapper but stops claiming to hold placeholder text", () => {
    const opened = open(
      P("Outside") +
        sdt(P("Inside"), `${PLACEHOLDER}<w:lock w:val="sdtLocked"/>`)
    );

    const after = type(opened, "Inside", "!");

    expect(blockTypes(after.doc)).toEqual(["paragraph", "sdtBlock"]);
    const written = exported(after, opened);
    expect(written).not.toContain("showingPlcHdr");
    expect(written).toContain('<w:lock w:val="sdtLocked"/>');
    expect(written).toContain("I!nside");
  });

  it("drops the flag off an inline control too", () => {
    const opened = open(
      `<w:p>${R("Plain")}${sdt(R("Inside"), `${PLACEHOLDER}<w:lock w:val="sdtLocked"/>`)}</w:p>`
    );

    const after = type(opened, "Inside", "!");

    const written = exported(after, opened);
    expect(written).toContain("<w:sdt>");
    expect(written).not.toContain("showingPlcHdr");
  });
});

describe("a control standing around a whole row", () => {
  const CELL = (text: string) => `<w:tc>${P(text)}</w:tc>`;
  const ROW = (...cells: string[]) => `<w:tr>${cells.join("")}</w:tr>`;

  const table = (firstRow: string) =>
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    firstRow +
    ROW(CELL("Under")) +
    "</w:tbl>";

  /** The row itself stays exactly where it stood; only what the control said about it goes */
  it("lifts the wrapper off the row on the first edit in one of its cells", () => {
    const opened = open(table(sdt(ROW(CELL("Inside")), TEMPORARY)));

    const after = type(opened, "Inside", "!");
    const row = after.doc.child(0).child(0);

    expect(row.type.name).toBe("tableRow");
    expect(row.attrs.sdtPrefix).toBeNull();
    expect(row.textContent).toBe("I!nside");
    expect(exported(after, opened)).not.toContain("<w:sdt>");
  });

  it("stands where the edit was in another row", () => {
    const opened = open(table(sdt(ROW(CELL("Inside")), TEMPORARY)));

    const after = type(opened, "Under", "!");

    expect(after.doc.child(0).child(0).attrs.sdtPrefix).not.toBeNull();
    expect(exported(after, opened)).toContain("<w:temporary/>");
  });

  it("drops the placeholder flag off the row and keeps the control", () => {
    const opened = open(table(sdt(ROW(CELL("Inside")), PLACEHOLDER)));

    const after = type(opened, "Inside", "!");
    const row = after.doc.child(0).child(0);

    expect(row.attrs.sdtShowingPlaceholder).toBe(false);
    expect(row.attrs.sdtPrefix).not.toContain("showingPlcHdr");
    const written = exported(after, opened);
    expect(written).toContain("<w:sdt>");
    expect(written).not.toContain("showingPlcHdr");
  });
});

describe("several controls settling in one round", () => {
  it("lifts a group control and a plain one from the same edit", () => {
    const opened = open(
      sdt(sdt(P("Grouped"), "", 2), `${TEMPORARY}<w:group/>`) +
        sdt(P("Plain"), TEMPORARY, 3)
    );
    const grouped = posOfText(opened.state.doc, "Grouped");
    const plain = posOfText(opened.state.doc, "Plain");
    const placed = select(opened.state, grouped);

    const after = placed.apply(
      placed.tr.insertText("!", plain).insertText("!", grouped)
    );

    expect(after.doc.textContent).toBe("G!roupedP!lain");
    const written = exported(after, opened);
    expect(written).not.toContain("<w:temporary/>");
    expect(written).not.toContain("<w:group/>");
  });
});

/** The package of `makeNotesDocx`, its footnote body wrapped in a control stating these props */
function notesWithControl(props: string): Uint8Array {
  const parts = unzipSync(makeNotesDocx());
  const body =
    "<w:p><w:r><w:footnoteRef/></w:r>" +
    '<w:r><w:t xml:space="preserve">Footnote body</w:t></w:r></w:p>';
  parts["word/footnotes.xml"] = new TextEncoder().encode(
    `<w:footnotes xmlns:w="${W_NS}">` +
      '<w:footnote w:id="-1" w:type="separator">' +
      "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
      `<w:footnote w:id="2">${sdt(body, props, 9)}</w:footnote>` +
      "</w:footnotes>"
  );
  return zipSync(parts);
}

/** The state the footnote of that package is edited in, built as the note surface builds one */
function footnoteState(props: string): EditorState {
  const opening = importDocx(notesWithControl(props));
  const main = editorStateForSession(opening);
  const story = storyOf(main.doc, storyKey("footnote", "2"));
  if (story === null) throw new Error("the document holds no footnote 2");
  const document = documentOf(main);
  return storyEditorState({
    story,
    document: storyDocument(document, document.geometry),
    protection: "none",
    keys: {},
    plugins: [],
    normalizers: [],
    takes: NO_CAPABILITY,
  });
}

describe("a control inside a side story", () => {
  it("settles in the footnote it stands in, as one in the body does", () => {
    const state = footnoteState(TEMPORARY);
    const at = posOfText(state.doc, "Footnote body");
    const placed = select(state, at);

    const after = placed.apply(placed.tr.insertText("!", at));

    expect(blockTypes(after.doc)).toEqual(["paragraph"]);
    expect(after.doc.textContent).toContain("F!ootnote body");
  });

  it("drops the placeholder flag there too", () => {
    const state = footnoteState(PLACEHOLDER);
    const at = posOfText(state.doc, "Footnote body");
    const placed = select(state, at);

    const after = placed.apply(placed.tr.insertText("!", at));

    const control = after.doc.child(0);
    expect(control.type.name).toBe("sdtBlock");
    expect(control.attrs.sdtPrefix).not.toContain("showingPlcHdr");
  });
});
