// @vitest-environment node
import { describe, expect, it } from "vitest";
import { tabSideAtX } from "./tabPointer";

describe("tabSideAtX", () => {
  it("reverses the visual halves for right-to-left text", () => {
    expect(tabSideAtX(100, 140, 110, "ltr")).toBe("before");
    expect(tabSideAtX(100, 140, 130, "ltr")).toBe("after");
    expect(tabSideAtX(100, 140, 110, "rtl")).toBe("after");
    expect(tabSideAtX(100, 140, 130, "rtl")).toBe("before");
  });
});
