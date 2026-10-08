// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { createEditorView, editorStateForSession } from "../createEditor";

const SHADED_DEFAULTS =
  '<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:rPr>';

const TEXT = '<w:r><w:t xml:space="preserve">Text</w:t></w:r>';

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function drawn(body: string, rPrDefault?: string): HTMLElement {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const view = createEditorView({
    mount,
    state: editorStateForSession(importDocx(makeDocx(body, rPrDefault))),
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view.dom;
}

/** The runs holding no characters drawn as a wide blank */
function wideBlanks(dom: HTMLElement): number {
  return dom.querySelectorAll(
    ".docx-editor-run[data-paints] > .docx-editor-empty-run"
  ).length;
}

describe("a run holding no characters", () => {
  it("is drawn wide where its own properties highlight it", () => {
    const body = `<w:p>${TEXT}<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t/></w:r></w:p>`;
    expect(wideBlanks(drawn(body))).toBe(1);
  });

  it("is drawn wide where its own properties shade it", () => {
    const body = `<w:p><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:rPr><w:t/></w:r></w:p>`;
    expect(wideBlanks(drawn(body))).toBe(1);
  });

  it("takes no room where only the formatting it inherits paints a background", () => {
    const body = `<w:p>${TEXT}<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r></w:p>`;
    const dom = drawn(body, SHADED_DEFAULTS);
    expect(dom.querySelector(".docx-editor-empty-run")).not.toBeNull();
    expect(dom.querySelectorAll("[data-paints]")).toHaveLength(0);
  });

  it("takes no room where its own fill is switched off", () => {
    const body = `<w:p><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:rPr><w:t/></w:r></w:p>`;
    expect(wideBlanks(drawn(body, SHADED_DEFAULTS))).toBe(0);
  });
});
