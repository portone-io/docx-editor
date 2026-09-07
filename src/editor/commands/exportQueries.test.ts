// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { fixtureNames, makeDocx, readFixture } from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { createEditorState, editorStateForSession } from "../createEditor";
import { canExport, documentExportProblems } from "./exportQueries";
import { toggleNumberedList } from "./listCommands";

const BODY = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>';

describe("canExport", () => {
  it.each(fixtureNames)("is true for %s freshly opened", (name) => {
    const state = editorStateForSession(importDocx(readFixture(name)));
    expect(canExport(state)).toBe(true);
    expect(documentExportProblems(state)).toEqual([]);
  });

  it("turns false when a list starts in a document without numbering.xml", () => {
    const opened = importDocx(makeDocx(BODY));
    // The built-in list command refuses where there is no numbering part, so the list is
    // started the way a plugin or a programmatic transform would start one
    const listed = runCommand(
      select(createEditorState(opened.doc), 1),
      toggleNumberedList
    );
    const state = editorStateForSession({
      doc: listed.doc,
      session: opened.session,
    });

    expect(canExport(state)).toBe(false);
    expect(documentExportProblems(state)).toEqual([
      {
        code: "missing-numbering-part",
        message:
          "cannot add a new list to a document that has no numbering.xml",
      },
    ]);
  });

  it("answers the same list without a second walk while the document stands", () => {
    const state = editorStateForSession(importDocx(makeDocx(BODY)));
    const moved = select(state, 2);

    expect(documentExportProblems(moved)).toBe(documentExportProblems(state));
  });
});

describe("a state built without a session", () => {
  it("reports no problem", () => {
    const { doc } = importDocx(makeDocx(BODY));
    const listed = runCommand(
      select(createEditorState(doc), 1),
      toggleNumberedList
    );

    expect(documentExportProblems(listed)).toEqual([]);
    expect(canExport(listed)).toBe(true);
  });
});
