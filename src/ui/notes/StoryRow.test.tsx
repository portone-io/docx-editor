// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { storyMarkup } from "../../editor/stories/storyMarkup";
import { docxSchema } from "../../schema";
import { DEFAULT_FONT_FALLBACKS } from "../../styles/fontStack";
import { StoryRow, type StoryRowProps } from "./StoryRow";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => host.remove());

/** A footnote's story as the file writes one: its own reference mark, then its text */
function footnoteStory(text: string) {
  const { doc, paragraph, rawRunContent } = docxSchema.nodes;
  return doc.create(null, [
    paragraph.create(null, [
      rawRunContent.create({ element: "footnoteRef", display: "chip" }),
      docxSchema.text(text),
    ]),
  ]);
}

describe("a note row", () => {
  it("does not rebuild the markup of a note whose story did not change", () => {
    const draw = vi.fn(storyMarkup);
    const first = footnoteStory("First");
    const second = footnoteStory("Second");
    const root = createRoot(host);
    const show = (
      props: Pick<StoryRowProps, "story" | "label" | "hidden">
    ): void =>
      act(() =>
        root.render(
          <StoryRow
            noteKey="footnote:1"
            fontFallbacks={DEFAULT_FONT_FALLBACKS}
            draw={draw}
            {...props}
          />
        )
      );

    show({ story: first, label: "1" });
    expect(draw).toHaveBeenCalledTimes(1);
    expect(host.textContent).toBe("1First");

    // Drawn again for something else, such as its height becoming known
    show({ story: first, label: "1", hidden: true });
    expect(draw).toHaveBeenCalledTimes(1);

    // Mounted again on another page, it takes the markup its story was drawn as
    act(() => root.render(null));
    show({ story: first, label: "1" });
    expect(draw).toHaveBeenCalledTimes(1);
    expect(host.textContent).toBe("1First");

    show({ story: second, label: "1" });
    expect(draw).toHaveBeenCalledTimes(2);
    expect(host.textContent).toBe("1Second");

    show({ story: second, label: "2" });
    expect(draw).toHaveBeenCalledTimes(3);
    expect(host.textContent).toBe("2Second");

    act(() => root.unmount());
  });
});
