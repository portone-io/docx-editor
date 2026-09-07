// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

/**
 * `EDIT_GUARDS` is built while `./guards` is evaluated, out of the guards their own modules export,
 * so the list is whole only for as long as no guard's module reads `./guards` back.
 *
 * This loads a guard's module before the list on purpose. Were the two to point at each other, the
 * bindings the list is built from would not be assigned yet when the module that holds them is
 * entered first, and the guards written elsewhere would land in the list as holes rather than
 * raising anything: the bundler lowers the declarations to `var`, so there is no temporal dead
 * zone to trip over, and the first edited transaction reads `change` off nothing.
 *
 * The imports are dynamic, and the registry is emptied first, because the suite shares one module
 * registry across files (`isolate: false`) and because a static import list is sorted by the
 * formatter, either of which would decide the order this answers for.
 */
describe("the guard list, reached through a guard's own module first", () => {
  it("holds every registered guard whole", async () => {
    vi.resetModules();
    await import("./preservedGuards");
    const { EDIT_GUARDS } = await import("./guards");

    expect(EDIT_GUARDS.map((guard) => guard.name)).toEqual([
      "protection",
      "lock",
      "bookmark",
      "note",
      "section",
    ]);
  });
});
