// @vitest-environment jsdom

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { posOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import { docxSchema } from "../schema/index";
import type { EditingProtection } from "../schema/protection";
import { editorStateForSession } from "./createEditor";
import { editRefusal } from "./editRefusal";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const cell = (inner: string) =>
  `<w:tc><w:tcPr><w:shd w:val="clear" w:fill="auto"/></w:tcPr><w:p>${inner}</w:p></w:tc>`;

const sdtPr = (tag: string, id: number, lock: string) =>
  `<w:sdtPr><w:alias w:val="${tag} alias"/><w:tag w:val="${tag}"/>` +
  `<w:id w:val="${id}"/><w:lock w:val="${lock}"/></w:sdtPr>`;

const inlineControl = (tag: string, id: number, lock: string, text: string) =>
  `<w:sdt>${sdtPr(tag, id, lock)}<w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;

/** A paragraph holding a locked field, and a table whose second row a locked control wraps */
const BODY =
  `<w:p>${run("Name: ")}${inlineControl("NAME", 11, "sdtContentLocked", "Kim")}${run(" end")}</w:p>` +
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cell(run("Head"))}${cell(run("Open"))}</w:tr>` +
  `<w:sdt>${sdtPr("PRICE_ROWS", -42, "sdtContentLocked")}<w:sdtContent>` +
  `<w:tr>${cell(run("Row"))}${cell(inlineControl("PRICE", 12, "contentLocked", "Fee"))}</w:tr>` +
  "</w:sdtContent></w:sdt>" +
  "</w:tbl>" +
  `<w:p>${run("Free text")}</w:p>`;

function opened(protection: EditingProtection = "none"): EditorState {
  return editorStateForSession(importDocx(makeDocx(BODY)), {
    protection,
    author: { id: "me", name: "Me" },
  });
}

function rowPos(doc: PMNode, index: number): number {
  let found = -1;
  let seen = 0;
  doc.descendants((node, pos) => {
    if (found >= 0) return false;
    if (node.type.spec.tableRole === "row") {
      if (seen === index) found = pos;
      seen += 1;
    }
    return true;
  });
  if (found < 0) throw new Error(`no row ${index}`);
  return found;
}

describe("editRefusal", () => {
  it("names the locked field a character typed into it was refused by", () => {
    const state = opened();
    const at = posOfText(state.doc, "Kim");
    const refusal = editRefusal(state.tr.insertText("x", at), state);

    expect(refusal).toEqual({
      reason: "lock",
      action: "insert",
      pos: at,
      controls: [
        {
          tag: "NAME",
          alias: "NAME alias",
          id: 11,
          lock: "sdtContentLocked",
          group: false,
          level: "inline",
          pos: at - 1,
        },
      ],
    });
  });

  it("names the row control around a cell whose text is deleted", () => {
    const state = opened();
    const at = posOfText(state.doc, "Row");
    const refusal = editRefusal(state.tr.delete(at, at + 1), state);

    expect(refusal?.reason).toBe("lock");
    expect(refusal?.action).toBe("delete");
    expect(refusal?.controls).toEqual([
      {
        tag: "PRICE_ROWS",
        alias: "PRICE_ROWS alias",
        id: -42,
        lock: "sdtContentLocked",
        group: false,
        level: "row",
        pos: rowPos(state.doc, 1),
      },
    ]);
  });

  it("names the innermost control first", () => {
    const state = opened();
    const at = posOfText(state.doc, "Fee");
    const refusal = editRefusal(state.tr.insertText("y", at, at + 1), state);

    expect(refusal?.action).toBe("replace");
    expect(refusal?.controls.map((control) => control.tag)).toEqual([
      "PRICE",
      "PRICE_ROWS",
    ]);
    expect(refusal?.controls[0]?.lock).toBe("contentLocked");
  });

  it("names the row a deletion of the whole row would take away", () => {
    const state = opened();
    const from = rowPos(state.doc, 1);
    const row = state.doc.nodeAt(from);
    if (!row) throw new Error("no row");
    const refusal = editRefusal(
      state.tr.delete(from, from + row.nodeSize),
      state
    );

    expect(refusal?.reason).toBe("lock");
    expect(refusal?.controls.map((control) => control.level)).toContain("row");
  });

  it("reports a formatting change as one", () => {
    const state = opened();
    const at = posOfText(state.doc, "Kim");
    const refusal = editRefusal(
      state.tr.addMark(
        at,
        at + 2,
        docxSchema.marks.link.create({ href: "https://example.com" })
      ),
      state
    );

    expect(refusal?.action).toBe("format");
    expect(refusal?.controls.map((control) => control.tag)).toEqual(["NAME"]);
  });

  it("names the protection, and no control, when the mode takes no body edit", () => {
    const state = opened("comments");
    const at = posOfText(state.doc, "Free text");
    const refusal = editRefusal(state.tr.insertText("z", at), state);

    expect(refusal).toEqual({
      reason: "protection",
      action: "insert",
      pos: at,
      controls: [],
    });
  });

  it("answers null for an edit the guards let through", () => {
    const state = opened();
    const at = posOfText(state.doc, "Free text");

    expect(editRefusal(state.tr.insertText("z", at), state)).toBeNull();
    expect(
      editRefusal(state.tr.insertText("z", posOfText(state.doc, "Open")), state)
    ).toBeNull();
  });
});
