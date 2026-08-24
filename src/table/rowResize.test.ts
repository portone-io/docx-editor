// @vitest-environment jsdom
import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  cell,
  firstTable,
  row,
  rowWith,
  schema,
  tableDoc,
} from "./__testing__/tables";
import {
  buildResizeRowTransaction,
  MIN_ROW_HEIGHT_PT,
  resizedRowHeight,
} from "./rowResize";

function resizeFirstRow(doc: ReturnType<typeof tableDoc>, heightPt: number) {
  const state = EditorState.create({ doc, schema });
  const { start } = firstTable(doc);
  const tr = buildResizeRowTransaction(state, {
    tablePos: start - 1,
    row: 0,
    heightPt,
  });
  return tr ? state.apply(tr) : null;
}

describe("resizedRowHeight", () => {
  it("converts screen movement to points at the current zoom", () => {
    expect(resizedRowHeight(20, 16, 1)).toBe(32);
    expect(resizedRowHeight(20, 8, 0.5)).toBe(32);
  });

  it("does not shrink below the draggable floor", () => {
    expect(resizedRowHeight(20, -100, 1)).toBe(MIN_ROW_HEIGHT_PT);
  });
});

describe("buildResizeRowTransaction", () => {
  it("writes the height on the row without changing its cells", () => {
    const before = tableDoc([row(cell("A"), cell("B"))]);
    const after = resizeFirstRow(before, 24);
    const table = after ? firstTable(after.doc).table : null;
    const resized = table?.child(0);

    expect(resized?.attrs.trPr).toBe(
      '<w:trPr><w:trHeight w:val="480" w:hRule="atLeast"/></w:trPr>'
    );
    expect(resized?.attrs.format).toEqual({
      height: { rule: "atLeast", pt: 24 },
    });
    expect(resized?.child(0)).toBe(firstTable(before).table.child(0).child(0));
  });

  it("preserves an exact rule already written on the row", () => {
    const before = tableDoc([
      rowWith(
        {
          trPr: '<w:trPr><w:trHeight w:val="400" w:hRule="exact"/></w:trPr>',
          format: { height: { rule: "exact", pt: 20 } },
        },
        cell("A")
      ),
    ]);
    const after = resizeFirstRow(before, 30);

    expect(
      after ? firstTable(after.doc).table.child(0).attrs : null
    ).toMatchObject({
      trPr: '<w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>',
      format: { height: { rule: "exact", pt: 30 } },
    });
  });

  it("does not resize a row containing a content-locked cell", () => {
    const before = tableDoc([row(cell("Locked", { sdtContentsLocked: true }))]);
    expect(resizeFirstRow(before, 24)).toBeNull();
  });
});
