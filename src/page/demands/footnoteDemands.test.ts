// @vitest-environment jsdom
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it, vi } from "vitest";
import { docxSchema } from "../../schema";
import type { MeasureTarget } from "../blockKinds";
import { FOOTNOTE_BAND, footnoteDemands } from "./footnoteDemands";

const { doc, paragraph, table, tableRow, tableCell, noteReference } =
  docxSchema.nodes;

function reference(kind: string, id: string) {
  return noteReference.create({ kind, id, label: id });
}

/**
 * A plain paragraph, one referring to two footnotes, one referring to an endnote, one referring
 * to a note of a kind the editor does not model, and a table
 */
function referringDocument() {
  return doc.create(null, [
    paragraph.create(null, [docxSchema.text("plain")]),
    paragraph.create(null, [
      docxSchema.text("a"),
      reference("footnote", "2"),
      docxSchema.text("b"),
      reference("footnote", "4"),
    ]),
    paragraph.create(null, [docxSchema.text("c"), reference("endnote", "3")]),
    // A kind this editor does not model, which the file may still carry
    paragraph.create(null, [docxSchema.text("d"), reference("sidenote", "5")]),
    table.create(null, [
      tableRow.create(null, [
        tableCell.create(null, [
          paragraph.create(null, [reference("footnote", "6")]),
        ]),
      ]),
    ]),
  ]);
}

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
});

function mounted(): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  view = new EditorView(mount, {
    state: EditorState.create({ doc: referringDocument() }),
  });
  return view;
}

/** A box standing at `top` on the screen, which is all a demand reads off a reference */
function boxAt(top: number): DOMRect {
  return {
    x: 0,
    y: top,
    top,
    left: 0,
    width: 8,
    height: 12,
    right: 8,
    bottom: top + 12,
    toJSON: () => ({}),
  };
}

/** The block at this index as the measurement hands it over, on a sheet read 10 below the screen */
function blockTarget(live: EditorView, index: number): MeasureTarget {
  let pos = 0;
  live.state.doc.forEach((_node, offset, at) => {
    if (at === index) pos = offset;
  });
  const node = live.state.doc.child(index);
  const dom = live.nodeDOM(pos);
  if (!(dom instanceof HTMLElement)) throw new Error("the block was not drawn");
  return {
    view: live,
    node,
    pos,
    dom,
    sheetY: (viewportY) => viewportY - 10,
    top: 100,
    scale: 1,
  };
}

describe("the room footnote references ask for", () => {
  it("asks for each footnote a block refers to where its reference is drawn", () => {
    const live = mounted();
    const tops = [130, 150];
    live.dom.querySelectorAll("sup").forEach((sup, index) => {
      const top = tops[index];
      if (top !== undefined) sup.getBoundingClientRect = () => boxAt(top);
    });

    expect(footnoteDemands.demandsIn(blockTarget(live, 1))).toEqual([
      { offset: 20, id: "footnote:2", band: FOOTNOTE_BAND },
      { offset: 40, id: "footnote:4", band: FOOTNOTE_BAND },
    ]);
  });

  it("asks for a footnote whose reference stands in a table cell", () => {
    const live = mounted();

    expect(
      footnoteDemands.demandsIn(blockTarget(live, 4)).map((demand) => demand.id)
    ).toEqual(["footnote:6"]);
  });

  it("keeps no room for a reference of a kind this editor does not model", () => {
    const live = mounted();

    // Claiming it would reserve the band's overhead and draw a rule over a strip holding nothing
    expect(footnoteDemands.demandsIn(blockTarget(live, 3))).toEqual([]);
  });

  it("answers a block holding no footnote reference without reading the page", () => {
    const live = mounted();
    const plain = blockTarget(live, 0);
    const endnote = blockTarget(live, 2);
    const unmodelled = blockTarget(live, 3);
    const nodeDOM = vi.spyOn(live, "nodeDOM");
    const sheetY = vi.fn((viewportY: number) => viewportY);
    const drawn = vi.spyOn(plain.dom, "getBoundingClientRect");

    expect(footnoteDemands.demandsIn({ ...plain, sheetY })).toEqual([]);
    expect(footnoteDemands.demandsIn({ ...endnote, sheetY })).toEqual([]);
    expect(footnoteDemands.demandsIn({ ...unmodelled, sheetY })).toEqual([]);
    expect(nodeDOM).not.toHaveBeenCalled();
    expect(sheetY).not.toHaveBeenCalled();
    expect(drawn).not.toHaveBeenCalled();
  });
});
