// @vitest-environment jsdom
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { editorAttributes } from "../styles/classNames";
import { setTableContinuations } from "./pageDecorations";
import { measureTable } from "./tableMeasurements";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  document.body.replaceChildren();
});

function paragraph(text: string) {
  return docxSchema.nodes.paragraph.create({}, [docxSchema.text(text)]);
}

function cell(text: string, rowspan = 1) {
  return docxSchema.nodes.tableCell.create({ rowspan }, [paragraph(text)]);
}

function mountedTable(): {
  live: EditorView;
  table: HTMLElement;
  rowPositions: number[];
} {
  const rows = [
    docxSchema.nodes.tableRow.create(
      { format: { repeatHeader: true, cantSplit: true } },
      [cell("Heading A"), cell("Heading B")]
    ),
    docxSchema.nodes.tableRow.create({}, [cell("Merged", 2), cell("B1")]),
    docxSchema.nodes.tableRow.create({}, [cell("B2")]),
    docxSchema.nodes.tableRow.create({}, [cell("A3"), cell("B3")]),
  ];
  const tableNode = docxSchema.nodes.table.create(
    { gridCols: [1000, 1000] },
    rows
  );
  const mount = document.createElement("div");
  document.body.append(mount);
  view = new EditorView(mount, {
    state: createEditorState(docxSchema.nodes.doc.create(null, [tableNode])),
  });
  const table = view.nodeDOM(0);
  if (!(table instanceof HTMLElement)) throw new Error("table DOM not found");

  const rowPositions: number[] = [];
  tableNode.forEach((_row, offset) => rowPositions.push(1 + offset));
  return { live: view, table, rowPositions };
}

function rect(element: Element, top: number, height: number): void {
  element.getBoundingClientRect = () => new DOMRect(0, top, 400, height);
}

function drawNaturalRows(
  live: EditorView,
  table: HTMLElement,
  rowPositions: readonly number[],
  scale = 1
): void {
  rect(table, 0, 340 * scale);
  const tops = [0, 40, 140, 240];
  const heights = [40, 100, 100, 100];
  rowPositions.forEach((pos, index) => {
    const row = live.nodeDOM(pos);
    if (row instanceof Element)
      rect(row, (tops[index] ?? 0) * scale, (heights[index] ?? 0) * scale);
  });
}

describe("measureTable", () => {
  it("repeats the leading header and excludes a boundary crossed by rowspan", () => {
    const { live, table, rowPositions } = mountedTable();
    drawNaturalRows(live, table, rowPositions);

    expect(measureTable(live, live.state.doc.child(0), 0, table)).toEqual({
      table: {
        boundaries: [
          { pos: rowPositions[1], offset: 40 },
          { pos: rowPositions[3], offset: 240 },
        ],
        firstPageMinimum: 240,
        repeatHeaderHeight: 40,
        headerRows: [rowPositions[0]],
        headerSignature: JSON.stringify([
          live.state.doc.child(0).child(0).toJSON(),
        ]),
        columns: 2,
      },
      appliedHeight: 0,
    });
  });

  it("normalizes row measurements taken from a visually scaled table", () => {
    const { live, table, rowPositions } = mountedTable();
    drawNaturalRows(live, table, rowPositions, 0.6);

    expect(measureTable(live, live.state.doc.child(0), 0, table, 0.6)).toEqual({
      table: {
        boundaries: [
          { pos: rowPositions[1], offset: 40 },
          { pos: rowPositions[3], offset: 240 },
        ],
        firstPageMinimum: 240,
        repeatHeaderHeight: 40,
        headerRows: [rowPositions[0]],
        headerSignature: JSON.stringify([
          live.state.doc.child(0).child(0).toJSON(),
        ]),
        columns: 2,
      },
      appliedHeight: 0,
    });
  });

  it("subtracts its own spacer and repeated header on the next pass", () => {
    const { live, table, rowPositions } = mountedTable();
    setTableContinuations(live, [
      {
        pos: rowPositions[3] ?? 0,
        height: 200,
        headerRows: [rowPositions[0] ?? 0],
        headerSignature: "heading-a",
        columns: 2,
      },
    ]);

    rect(table, 0, 580);
    const originalTops = [0, 40, 140, 480];
    const originalHeights = [40, 100, 100, 100];
    rowPositions.forEach((pos, index) => {
      const row = live.nodeDOM(pos);
      if (row instanceof Element) {
        rect(row, originalTops[index] ?? 0, originalHeights[index] ?? 0);
      }
    });
    const spacer = table.querySelector(`[${editorAttributes.tablePageSpace}]`);
    const repeated = table.querySelector(
      `[${editorAttributes.tableRepeatedHeader}]`
    );
    if (!spacer || !repeated) throw new Error("pagination rows not found");
    rect(spacer, 240, 200);
    rect(repeated, 440, 40);

    const measured = measureTable(live, live.state.doc.child(0), 0, table);
    expect(measured?.appliedHeight).toBe(240);
    expect(measured?.table.boundaries).toEqual([
      { pos: rowPositions[1], offset: 40 },
      { pos: rowPositions[3], offset: 240 },
    ]);
    expect(measured?.table.firstPageMinimum).toBe(240);
  });
});
