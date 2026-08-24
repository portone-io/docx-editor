import { describe, expect, it } from "vitest";
import { layerTabStopDirectives, mergeTabStops } from "./tabStops";

describe("effective custom tab stops", () => {
  it("sorts positions and lets the later layer replace the same position", () => {
    expect(
      mergeTabStops(
        [
          { positionPt: 72, align: "center" },
          { positionPt: 36, align: "start", leader: "dot" },
        ],
        [{ positionPt: 72, align: "end" }]
      )
    ).toEqual([
      { positionPt: 36, align: "start", leader: "dot" },
      { positionPt: 72, align: "end" },
    ]);
  });

  it("removes only an inherited stop at the clear position", () => {
    expect(
      mergeTabStops(
        [
          { positionPt: 36, align: "start" },
          { positionPt: 72, align: "center" },
        ],
        [{ positionPt: 36, align: "clear" }]
      )
    ).toEqual([{ positionPt: 72, align: "center" }]);
  });

  it("retains a clear while intermediate layers are composed", () => {
    expect(
      layerTabStopDirectives([], [{ positionPt: 36, align: "clear" }])
    ).toEqual([{ positionPt: 36, align: "clear" }]);
  });
});
