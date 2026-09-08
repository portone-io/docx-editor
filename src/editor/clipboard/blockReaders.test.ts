// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { editorClassNames } from "../../styles/classNames";
import { INTERNAL_TOKEN_ATTRIBUTE } from "./internalChannel";
import { detectHtmlSource } from "./source";

const htmlDir = join(
  dirname(fileURLToPath(import.meta.url)),
  "__testing__",
  "html"
);

/** One of the clipboard shapes under `__testing__/html`, as the browser would hand it over */
function fixture(name: string): string {
  return readFileSync(join(htmlDir, name), "utf8");
}

function markup(html: string): DocumentFragment {
  const template = document.createElement("template");
  template.innerHTML = html;
  return template.content;
}

describe("recognizing where pasted HTML came from", () => {
  it("identifies the source from the captured HTML", () => {
    expect(detectHtmlSource(markup(fixture("word-list-table.html")))).toBe(
      "word"
    );
    expect(
      detectHtmlSource(markup(fixture("google-docs-formatting.html")))
    ).toBe("google-docs");
    expect(
      detectHtmlSource(markup(fixture("libreoffice-paragraphs.html")))
    ).toBe("libreoffice");
    expect(
      detectHtmlSource(
        markup(
          `<p class="${editorClassNames.paragraph}" ${INTERNAL_TOKEN_ATTRIBUTE}="copy-1">own</p>`
        )
      )
    ).toBe("editor");
    expect(detectHtmlSource(markup("<p>plain</p>"))).toBe("unknown");
  });
});
