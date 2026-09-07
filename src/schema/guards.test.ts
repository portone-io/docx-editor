// @vitest-environment jsdom

import { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { AddMarkStep } from "prosemirror-transform";
import { describe, expect, it, vi } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "../editor/createEditor";
import { displayOnly } from "./displayDerivation";
import {
  editShut,
  guardedCommand,
  openStretches,
  transactionAllowed,
} from "./guards";
import { docxSchema } from "./index";
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

/** A one row table whose left cell stands inside a control that shuts its contents */
const LOCKED_CELL_TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  "<w:tr>" +
  `<w:sdt>${LOCKED_PR}<w:sdtContent><w:tc><w:p>${runXml("Shut")}</w:p></w:tc></w:sdtContent></w:sdt>` +
  `<w:tc><w:p>${runXml("Open")}</w:p></w:tc>` +
  "</w:tr>" +
  "</w:tbl>";

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

  /**
   * A block rewritten around its content - given an alignment or an indent - is the one intent the
   * lock answers by the cell alone: a paragraph merely holding a locked control keeps both, and
   * only what a locked cell holds is shut (`editor/paragraphEdits`).
   */
  it("shuts a block intent inside a locked cell and leaves the cell beside it open", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const shut = textRange(state.doc, "Shut");
    const open = textRange(state.doc, "Open");

    expect(editShut(state, { kind: "block", at: shut.from })).toBe(true);
    expect(editShut(state, { kind: "block", at: open.from })).toBe(false);
  });
});

describe("guardedCommand", () => {
  /**
   * The command is both what a control is drawn from and what the click runs, so the two have to be
   * the same reading. One reporting yes to the button and then being refused at dispatch would draw
   * a live control that swallows the click.
   */
  it("answers the same with and without dispatch", () => {
    const state = opened(LOCKED_P);
    const stretches = [
      { text: "bc", allowed: false },
      { text: "a", allowed: true },
    ];

    for (const { text, allowed } of stretches) {
      const { from, to } = textRange(state.doc, text);
      const command = guardedCommand((current) => current.tr.delete(from, to));
      const dispatched: Transaction[] = [];

      expect(command(state), `asked about "${text}"`).toBe(allowed);
      expect(
        command(state, (tr) => dispatched.push(tr)),
        `run over "${text}"`
      ).toBe(allowed);
      expect(dispatched).toHaveLength(allowed ? 1 : 0);
    }
  });
});

describe("openStretches", () => {
  it("leaves the locked stretch out and keeps the rest", () => {
    const state = opened(LOCKED_P);
    const stretches = ["a", "bc", "d"].map((text) => ({
      text,
      ...textRange(state.doc, text),
    }));

    expect(
      openStretches(state, stretches, "mark").map((of) => of.text)
    ).toEqual(["a", "d"]);
    // The whole list is asked of each stretch, so a protection shutting the body leaves none open
    expect(
      openStretches(opened(LOCKED_P, "readOnly"), stretches, "mark")
    ).toEqual([]);
  });
});

describe("display mark updates under protection", () => {
  it.each(["readOnly", "comments"] as const)(
    "allows derived values but refuses source changes under %s",
    (protection) => {
      const state = opened(LOCKED_P, protection);
      const { from, to } = textRange(state.doc, "bc");
      const mark = docxSchema.marks.run.isInSet(
        state.doc.nodeAt(from)?.marks ?? []
      );
      if (!mark) throw new Error("expected imported run mark");
      const derived = state.tr
        .step(
          new AddMarkStep(
            from,
            to,
            mark.type.create({ ...mark.attrs, format: { bold: true } })
          )
        )
        .setMeta(displayOnly, true);
      expect(transactionAllowed(derived, state)).toBe(true);
      expect(state.applyTransaction(derived).transactions[0]).toBe(derived);
      const source = state.tr
        .step(
          new AddMarkStep(
            from,
            to,
            mark.type.create({
              ...mark.attrs,
              rPr: "<w:rPr><w:b/></w:rPr>",
              format: { bold: true },
            })
          )
        )
        .setMeta(displayOnly, true);
      expect(transactionAllowed(source, state)).toBe(false);
      expect(state.apply(source).doc).toBe(state.doc);
    }
  );
});
