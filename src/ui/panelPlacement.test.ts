// @vitest-environment node
/**
 * Where a floating panel lands.
 *
 * The point of these rules is that a panel is always whole and always on screen, whatever
 * the host app wraps the editor in, so what is checked here is every way a panel can run
 * out of room.
 */

import { describe, expect, it } from "vitest";
import { positionAtPoint, positionUnderControl } from "./panelPlacement";

const SCREEN = { width: 1000, height: 800 };
const PANEL = { width: 200, height: 150 };

describe("positionUnderControl", () => {
  it("hangs the panel under the control, lined up with its left edge", () => {
    expect(
      positionUnderControl({ left: 100, top: 40, bottom: 68 }, PANEL, SCREEN)
    ).toEqual({ left: 100, top: 72 });
  });

  it("slides the panel back in when it would run past the right edge", () => {
    expect(
      positionUnderControl({ left: 900, top: 40, bottom: 68 }, PANEL, SCREEN)
    ).toEqual({ left: 792, top: 72 });
  });

  it("pins a panel wider than the screen to the left edge, so its head stays visible", () => {
    expect(
      positionUnderControl(
        { left: 10, top: 40, bottom: 68 },
        { width: 1200, height: 150 },
        SCREEN
      )
    ).toEqual({ left: 8, top: 72 });
  });

  it("flips the panel above the control when there is no room under it", () => {
    expect(
      positionUnderControl({ left: 100, top: 500, bottom: 528 }, PANEL, {
        width: 1000,
        height: 560,
      })
    ).toEqual({ left: 100, top: 346 });
  });

  it("slides the panel up from the bottom edge when it fits neither under nor over the control", () => {
    expect(
      positionUnderControl(
        { left: 100, top: 100, bottom: 128 },
        { width: 200, height: 300 },
        { width: 1000, height: 340 }
      )
    ).toEqual({ left: 100, top: 32 });
  });

  it("pins a panel taller than the screen to the top edge", () => {
    expect(
      positionUnderControl(
        { left: 100, top: 100, bottom: 128 },
        { width: 200, height: 300 },
        { width: 1000, height: 200 }
      )
    ).toEqual({ left: 100, top: 8 });
  });
});

describe("positionAtPoint", () => {
  it("starts the panel at the point that opened it", () => {
    expect(
      positionAtPoint({ clientX: 120, clientY: 240 }, PANEL, SCREEN)
    ).toEqual({ left: 120, top: 240 });
  });

  it("pulls the panel in from the right and bottom edges", () => {
    expect(
      positionAtPoint(
        { clientX: 980, clientY: 700 },
        { width: 180, height: 280 },
        SCREEN
      )
    ).toEqual({ left: 812, top: 512 });
  });
});
