// @vitest-environment jsdom
import { DOMSerializer } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { docxSchema, runMarkSpec } from "../../schema";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../../styles/fontStack";
import { storyMarkup } from "./storyMarkup";

const { doc, paragraph, hardBreak } = docxSchema.nodes;

/** The style an editor view holding these fallbacks puts on a run (`editor/views/runMarkView`) */
function editorRunStyle(
  attrs: Record<string, unknown>,
  fontFallbacks: FontFallbacks
): string | null {
  const { dom } = DOMSerializer.renderSpec(
    document,
    runMarkSpec(attrs, fontFallbacks)
  );
  return dom instanceof Element ? dom.getAttribute("style") : null;
}

describe("a story drawn as markup", () => {
  it("draws a run the way an editor holding the same fallback fonts draws it", () => {
    const fontFallbacks: FontFallbacks = {
      ...DEFAULT_FONT_FALLBACKS,
      defaultStack: "Stand In, serif",
    };
    const run = docxSchema.marks.run.create({
      format: { fontFamily: '"Nowhere Office Font"', bold: true },
    });
    const story = doc.create(null, [
      paragraph.create(null, [docxSchema.text("words", [run])]),
    ]);

    const drawn = storyMarkup(story, { fontFallbacks }).querySelector("p span");

    expect(drawn?.getAttribute("style")).toBe(
      editorRunStyle(run.attrs, fontFallbacks)
    );
    expect(drawn?.getAttribute("style")).toContain("Stand In");
  });

  it("draws a node type through the spec handed in and every other one as the schema does", () => {
    const story = doc.create(null, [
      paragraph.create(null, [
        docxSchema.text("one"),
        hardBreak.create(),
        docxSchema.text("two"),
      ]),
      paragraph.create(null, [docxSchema.text("three")]),
    ]);

    const markup = storyMarkup(story, {
      fontFallbacks: DEFAULT_FONT_FALLBACKS,
      nodeSpecs: { hardBreak: () => ["span", { class: "stood-in" }, "/"] },
    });

    expect(
      [...markup.querySelectorAll("p")].map((drawn) => drawn.textContent)
    ).toEqual(["one/two", "three"]);
    expect(markup.querySelector("br")).toBeNull();
  });
});
