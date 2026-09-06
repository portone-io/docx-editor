// @vitest-environment jsdom
import { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it, vi } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "../editor/createEditor";
import { editShut, transactionAllowed } from "./guards";
import { historyReplay, unlockAllowed } from "./locks";
import type { EditingProtection } from "./protection";

const runXml = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

/** One paragraph reading "a", a control that shuts its contents over "bc", then "d" */
const LOCKED_P =
  `<w:p>${runXml("a")}` +
  `<w:sdt>${LOCKED_PR}<w:sdtContent>${runXml("bc")}</w:sdtContent></w:sdt>` +
  `${runXml("d")}</w:p>`;

/** One paragraph reading "mkn" with a bookmark range anchored around the "k" */
const BOOKMARK_P =
  "<w:p>" +
  runXml("m") +
  '<w:bookmarkStart w:id="9" w:name="b"/>' +
  runXml("k") +
  '<w:bookmarkEnd w:id="9"/>' +
  runXml("n") +
  "</w:p>";

function opened(
  body: string,
  protection: EditingProtection = "none"
): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc, {
    protection,
    author: { id: "me", name: "Me" },
  });
}

/** The stretch the first text node reading exactly this covers */
function textRange(doc: PMNode, needle: string): { from: number; to: number } {
  const found: { from: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (found.length === 0 && node.isText && node.text === needle) {
      found.push({ from: pos, to: pos + node.nodeSize });
    }
    return found.length === 0;
  });
  const first = found[0];
  if (first === undefined) throw new Error(`text not found: ${needle}`);
  return first;
}

/** The stretch the first node of this type covers */
function nodeRange(
  doc: PMNode,
  typeName: string
): { from: number; to: number } {
  const found: { from: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (found.length === 0 && node.type.name === typeName) {
      found.push({ from: pos, to: pos + node.nodeSize });
    }
    return found.length === 0;
  });
  const first = found[0];
  if (first === undefined) throw new Error(`no ${typeName} in the document`);
  return first;
}

describe("the guard list", () => {
  it("judges the protection before the passes and without them", () => {
    const state = opened(LOCKED_P, "readOnly");

    for (const pass of [unlockAllowed, historyReplay]) {
      const tr = state.tr.setMeta(pass, true).insertText("x", 1);
      expect(transactionAllowed(tr, state)).toBe(false);
    }
  });

  it("lets a pass lift the lock guard and no other", () => {
    const locked = opened(LOCKED_P);
    const control = textRange(locked.doc, "bc");
    expect(
      transactionAllowed(
        locked.tr
          .setMeta(unlockAllowed, true)
          .delete(control.from + 1, control.to),
        locked
      )
    ).toBe(true);

    const bookmarked = opened(BOOKMARK_P);
    const marker = nodeRange(bookmarked.doc, "rawInline");
    for (const pass of [unlockAllowed, historyReplay]) {
      const tr = bookmarked.tr
        .setMeta(pass, true)
        .delete(marker.from, marker.to);
      expect(transactionAllowed(tr, bookmarked)).toBe(false);
    }
  });

  /**
   * A guard that walks both documents does so only once a step has reached one of the nodes it
   * answers for, so a guard ahead of it in the list having already said no is visible as the walk
   * that never happens.
   */
  it("asks every guard and stops at the first refusal", () => {
    const open = opened(BOOKMARK_P);
    const marker = nodeRange(open.doc, "rawInline");
    expect(
      transactionAllowed(open.tr.delete(marker.from, marker.to), open)
    ).toBe(false);

    const shut = opened(BOOKMARK_P, "readOnly");
    const walked = vi.spyOn(PMNode.prototype, "descendants");
    try {
      expect(
        transactionAllowed(shut.tr.delete(marker.from, marker.to), shut)
      ).toBe(false);
      expect(walked).not.toHaveBeenCalled();
    } finally {
      walked.mockRestore();
    }
  });

  it("answers a transaction that changes no document with yes", () => {
    const state = opened(LOCKED_P, "readOnly");
    expect(transactionAllowed(state.tr.scrollIntoView(), state)).toBe(true);
  });
});

describe("editShut", () => {
  it("answers the protection, the locks and the preserved markers for one intent", () => {
    const locked = opened(LOCKED_P);
    const control = textRange(locked.doc, "bc");
    expect(editShut(locked, { kind: "insert", at: control.from + 1 })).toBe(
      true
    );
    expect(editShut(locked, { kind: "mark", ...control })).toBe(true);
    expect(editShut(locked, { kind: "insert", at: 1 })).toBe(false);

    const bookmarked = opened(BOOKMARK_P);
    const marker = nodeRange(bookmarked.doc, "rawInline");
    expect(editShut(bookmarked, { kind: "replace", ...marker })).toBe(true);
    expect(editShut(bookmarked, { kind: "insert", at: marker.to })).toBe(false);

    const readOnly = opened(BOOKMARK_P, "readOnly");
    expect(editShut(readOnly, { kind: "insert", at: 1 })).toBe(true);
    expect(editShut(readOnly, { kind: "block", at: 1 })).toBe(true);
  });
});
