// @vitest-environment jsdom
import { act, type ReactElement, useState } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { usePageRoom } from "./usePageRoom";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let root: Root | null = null;
let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  if (root) act(() => root?.unmount());
  root = null;
  host.remove();
});

const DRAWN_HEIGHT = 420;

/**
 * There is no layout here, so the layer answers with a height of its own. It is the height the
 * layer is *drawn* at, the scale already in it, which is what the hook writes.
 */
function drawn(node: HTMLElement | null): void {
  if (!node) return;
  node.getBoundingClientRect = () => new DOMRect(0, 0, 0, DRAWN_HEIGHT);
}

/**
 * A surface that draws neither the box nor the layer until it is told to, which is what one
 * refusing a file and then opening another does.
 */
function Surface({ scale }: { scale: number }): ReactElement {
  const [box, setBox] = useState<HTMLDivElement | null>(null);
  const [layer, setLayer] = useState<HTMLDivElement | null>(null);
  const [open, setOpen] = useState(false);
  usePageRoom(box, layer, scale);

  return (
    <>
      <button type="button" onClick={() => setOpen(true)}>
        open
      </button>
      {open && (
        <div data-testid="box" ref={setBox}>
          <div
            ref={(node) => {
              drawn(node);
              setLayer(node);
            }}
          />
        </div>
      )}
    </>
  );
}

function render(scale = 1): void {
  root = createRoot(host);
  act(() => root?.render(<Surface scale={scale} />));
}

function roomHeld(): string | undefined {
  return host.querySelector<HTMLElement>('[data-testid="box"]')?.style.height;
}

function openTheDocument(): void {
  const open = host.querySelector("button");
  if (!open) throw new Error("nothing to open the document with");
  act(() => open.click());
}

describe("usePageRoom", () => {
  it("holds the room the paper takes the moment it is drawn", () => {
    render();
    expect(roomHeld()).toBeUndefined();

    openTheDocument();

    expect(roomHeld()).toBe(`${DRAWN_HEIGHT}px`);
  });

  it("holds the height the paper is drawn at, the scale already in it", () => {
    render(0.5);
    openTheDocument();

    expect(roomHeld()).toBe(`${DRAWN_HEIGHT}px`);
  });
});
