// @vitest-environment jsdom
import { DOMSerializer } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { docxSchema, runMarkSpec } from "../../schema";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../../styles/fontStack";
import { noteNodeSpecs } from "../notes/noteSurface";
import { storyMarkup } from "./storyMarkup";

const { doc, paragraph, hardBreak, rawInline, rawRunContent, noteReference } =
  docxSchema.nodes;

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

  it("draws none of the document's own source, keeping only the attributes the stylesheet reads", () => {
    const run = docxSchema.marks.run.create({
      rPr: '<w:rPr><w:b/><w:u w:val="single"/></w:rPr>',
      rAttrs: 'w:rsidR="00A1B2C3"',
      format: { bold: true, underline: "single" },
    });
    const story = doc.create(null, [
      paragraph.create(
        {
          srcId: "p1",
          pAttrs: 'w14:paraId="1A2B3C4D"',
          pPr: '<w:pPr><w:pStyle w:val="FootnoteText"/><w:jc w:val="center"/></w:pPr>',
          format: { align: "center" },
        },
        [
          rawRunContent.create({
            xml: "<w:footnoteRef/>",
            element: "footnoteRef",
            display: "chip",
          }),
          rawInline.create({
            xml: '<w:bookmarkStart w:id="0" w:name="kept"/>',
            element: "bookmarkStart",
            display: "hidden",
          }),
          docxSchema.text("styled words", [run]),
          rawRunContent.create({
            xml: '<w:fldChar w:fldCharType="begin"/>',
            element: "fldChar",
            display: "chip",
            guarded: true,
          }),
          noteReference.create({
            kind: "footnote",
            id: "2",
            label: "2",
            referenceXml: '<w:footnoteReference w:id="2"/>',
          }),
        ]
      ),
    ]);

    const box = document.createElement("div");
    box.append(
      storyMarkup(story, {
        fontFallbacks: DEFAULT_FONT_FALLBACKS,
        nodeSpecs: noteNodeSpecs(() => "1"),
      })
    );

    const dataAttributes = new Set(
      [...box.querySelectorAll("*")].flatMap((element) =>
        element.getAttributeNames().filter((name) => name.startsWith("data-"))
      )
    );
    expect([...dataAttributes].sort()).toEqual([
      "data-display",
      "data-underline",
    ]);
    expect(box.innerHTML).not.toContain("<w:");
    expect(box.textContent).toContain("1styled words");
  });
});
