// @vitest-environment node
import { unzipSync, zipSync } from "fflate";
import { JSDOM } from "jsdom";
import { TextSelection } from "prosemirror-state";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { LETTER_SECT_PR, makeDocx } from "./__testing__/docx";
import { rangeOfText } from "./__testing__/editing";
import { exportDocx, importDocx, onlyCommentsChangedBy } from "./core";
import { addComment, updateComment } from "./editor/commands/commentCommands";
import { createEditorState } from "./editor/createEditor";

/**
 * The entry read where a server reads it, which is the one place the global it was told to
 * install is the only one there is.
 *
 * `site/content/docs/core.mdx` offers a `DOMParser` global as the alternative to the `xmlParser`
 * option and asks for nothing else, and the rest of the suite runs under jsdom, where every other
 * global is there to be reached for by accident.
 */
describe("under the one global the documentation offers a server", () => {
  // `vitest.config.ts` does not isolate files, so a global left behind is one the next node
  // environment file in this worker would find without having asked for it
  beforeAll(() => {
    const globals = globalThis as unknown as Record<string, unknown>;
    globals.DOMParser = new JSDOM().window.DOMParser;
  });

  afterAll(() => {
    const globals = globalThis as unknown as Record<string, unknown>;
    globals.DOMParser = undefined;
  });

  const author = { id: "me", name: "Me" };

  function commented(): { bytes: Uint8Array; commented: Uint8Array } {
    const parts = unzipSync(
      makeDocx(
        '<w:p><w:r><w:t xml:space="preserve">Alpha beta</w:t></w:r></w:p>' +
          LETTER_SECT_PR,
        '<w:sz w:val="20"/>'
      )
    );
    parts["[Content_Types].xml"] = new TextEncoder().encode(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>"
    );
    const bytes = zipSync(parts);
    const { doc, session } = importDocx(bytes);
    let state = createEditorState(doc, { author });
    const { from, to } = rangeOfText(state.doc, "beta");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to))
    );
    addComment({ text: "note", author: "Me", authorId: "me" })(
      state,
      (tr) => (state = state.apply(tr))
    );
    return { bytes, commented: exportDocx(state.doc, session) };
  }

  it("has none of the globals it was not promised", () => {
    const globals = globalThis as unknown as Record<string, unknown>;
    expect(globals.Node).toBeUndefined();
    expect(globals.Element).toBeUndefined();
    expect(globals.document).toBeUndefined();
  });

  it("holds for a comment this author added", () => {
    const { bytes, commented: submitted } = commented();
    expect(onlyCommentsChangedBy(bytes, submitted, "me")).toEqual({ ok: true });
  });

  it("holds for a comment this author then rewrote", () => {
    const { commented: before } = commented();
    const { doc, session } = importDocx(before);
    let state = createEditorState(doc, { author });
    expect(
      updateComment("0", "rewritten")(state, (tr) => (state = state.apply(tr)))
    ).toBe(true);

    expect(
      onlyCommentsChangedBy(before, exportDocx(state.doc, session), "me")
    ).toEqual({ ok: true });
  });

  it("does not hold for a body edit", () => {
    const { bytes, commented: submitted } = commented();
    const { doc, session } = importDocx(submitted);
    const edited = doc.copy(
      doc.content.replaceChild(
        0,
        doc.child(0).type.create(doc.child(0).attrs, [])
      )
    );

    expect(
      onlyCommentsChangedBy(bytes, exportDocx(edited, session), "me")
    ).toEqual({
      ok: false,
      reason: "body-changed",
    });
  });
});
