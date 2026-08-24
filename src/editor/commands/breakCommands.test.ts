// @vitest-environment jsdom
import type { Command, EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { posOfText, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { createEditorState } from "../createEditor";
import { insertLineBreak, insertPageBreak } from "./breakCommands";

/** A paragraph whose middle stretch stands in a control locked against deletion alone */
const BODY =
  "<w:p><w:r><w:t>open</w:t></w:r>" +
  '<w:sdt><w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtLocked"/></w:sdtPr>' +
  "<w:sdtContent><w:r><w:t>LOCKED</w:t></w:r></w:sdtContent></w:sdt>" +
  "<w:r><w:t>tail</w:t></w:r></w:p>";

function opened(): EditorState {
  return createEditorState(importDocx(makeDocx(BODY)).doc);
}

/** The selection covering exactly this text, which is the whole of the control wrapping it */
function over(needle: string): EditorState {
  const state = opened();
  const inside = posOfText(state.doc, needle);
  return select(state, inside - 1, inside - 1 + needle.length);
}

/** What the command answered, and the state its transactions left behind */
function attempt(
  state: EditorState,
  command: Command
): { answered: boolean; after: EditorState } {
  let after = state;
  const answered = command(state, (tr) => {
    after = after.apply(tr);
  });
  return { answered, after };
}

const COMMANDS: ReadonlyArray<[name: string, command: Command]> = [
  ["insertLineBreak", insertLineBreak],
  ["insertPageBreak", insertPageBreak],
];

/**
 * The break goes in in place of whatever is selected, so the deletion clause answers for it
 * (`schema/locks`): a control locked against deletion alone refuses it even though editing inside
 * it would have gone through. Asked with the contents clause instead, both commands reported that
 * the break had gone in and the guard then turned the transaction down.
 */
describe.each(COMMANDS)("%s over a locked control", (_name, command) => {
  it("says nothing goes in over a control locked against deletion alone", () => {
    const before = over("LOCKED");
    const { answered, after } = attempt(before, command);
    expect(answered).toBe(false);
    expect(after.doc.eq(before.doc)).toBe(true);
  });

  it("goes in over a selection of the same shape no lock covers", () => {
    const before = over("tail");
    const { answered, after } = attempt(before, command);
    expect(answered).toBe(true);
    expect(after.doc.eq(before.doc)).toBe(false);
  });
});
