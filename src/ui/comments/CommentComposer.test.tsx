// @vitest-environment jsdom
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { renderInto } from "../../__testing__/react";
import type { CommentAuthor } from "../../editor/commands/commentCommands";
import { CommentComposer } from "./CommentComposer";

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

const GRACE: CommentAuthor = { id: "grace", name: "Grace" };

function field(): HTMLTextAreaElement {
  const found = host.querySelector("textarea");
  if (!(found instanceof HTMLTextAreaElement)) {
    throw new Error("the composer is not open");
  }
  return found;
}

function write(value: string): void {
  const set = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  if (!set) throw new Error("no value setter on HTMLTextAreaElement");
  const input = field();
  act(() => {
    set.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

function submit(): void {
  const button = host.querySelector('button[type="submit"]');
  if (!(button instanceof HTMLButtonElement)) {
    throw new Error("the composer has no submit button");
  }
  act(() => button.click());
}

const said = () => host.querySelector('[role="alert"]')?.textContent ?? null;

/** Waits out one animation frame and whatever React does with what it left */
function frame(): Promise<void> {
  return act(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => resolve());
      })
  );
}

interface Attempt {
  text: string;
}

function mount(answer: (text: string) => string | null) {
  const asked: Attempt[] = [];
  let closed = 0;
  const unmount = renderInto(
    host,
    <CommentComposer
      author={GRACE}
      label="Comment text"
      submitLabel="Comment"
      onClose={() => {
        closed += 1;
      }}
      onSubmit={(text) => {
        asked.push({ text });
        return answer(text);
      }}
    />
  );
  return { asked, closed: () => closed, unmount };
}

describe("the comment composer form", () => {
  it("hands the text over and closes once it is written", () => {
    const { asked, closed, unmount } = mount(() => null);
    write("Looks wrong");
    submit();

    expect(asked).toEqual([{ text: "Looks wrong" }]);
    expect(closed()).toBe(1);
    expect(said()).toBeNull();
    unmount();
  });

  it("keeps the text and says why where the comment was refused", () => {
    const { closed, unmount } = mount(() => "The text it marked is locked.");
    write("Looks wrong");
    submit();

    // Standing there answering nothing is what would lose the reader's words
    expect(said()).toBe("The text it marked is locked.");
    expect(closed()).toBe(0);
    expect(field().value).toBe("Looks wrong");

    // Writing again takes the answer away, so it is never left standing over a newer attempt
    write("Looks wrong here");
    expect(said()).toBeNull();
    unmount();
  });

  /**
   * The scroll itself is a browser matter and `e2e/comments.spec.ts` holds it there. What is held
   * here is the pair of calls that keep it right: focusing an off-screen form without letting the
   * browser scroll to it, and asking for it to be shown only as far as it has to be. jsdom has no
   * `scrollIntoView` at all, so both are recorded off the prototype.
   */
  it("takes the focus without taking the page with it", async () => {
    const focused: unknown[] = [];
    const shown: unknown[] = [];
    const realFocus = HTMLTextAreaElement.prototype.focus;
    HTMLTextAreaElement.prototype.focus = function record(
      this: HTMLTextAreaElement,
      options?: FocusOptions
    ) {
      focused.push(options);
      realFocus.call(this, options);
    };
    Element.prototype.scrollIntoView = function record(
      options?: boolean | ScrollIntoViewOptions
    ) {
      shown.push(options);
    };
    try {
      const { unmount } = mount(() => null);
      expect(focused).toEqual([{ preventScroll: true }]);
      // The rail places the form on the frame after the one it was drawn in, so the asking to be
      // shown waits for that frame rather than pointing at where the form stood before
      expect(shown).toEqual([]);
      await frame();
      expect(shown).toEqual([{ block: "nearest" }]);
      unmount();
    } finally {
      HTMLTextAreaElement.prototype.focus = realFocus;
      Reflect.deleteProperty(Element.prototype, "scrollIntoView");
    }
  });
});
