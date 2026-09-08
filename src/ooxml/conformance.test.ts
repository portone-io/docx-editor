// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { bindsWritingPrefix, conformanceOf, STRICT_W_NS } from "./conformance";
import { W_NS } from "./names";
import { parseXml } from "./xml";

const rootOf = (xml: string): Element => parseXml(xml).documentElement;

describe("conformanceOf", () => {
  it("names a Strict main part by its namespace", () => {
    expect(
      conformanceOf(rootOf(`<w:document xmlns:w="${STRICT_W_NS}"/>`))
    ).toBe("strict");
  });

  it("names a Transitional main part by its namespace", () => {
    expect(conformanceOf(rootOf(`<w:document xmlns:w="${W_NS}"/>`))).toBe(
      "transitional"
    );
  });

  it("reads the namespace rather than the prefix in front of it", () => {
    expect(
      conformanceOf(rootOf(`<p:document xmlns:p="${STRICT_W_NS}"/>`))
    ).toBe("strict");
    expect(conformanceOf(rootOf(`<document xmlns="${W_NS}"/>`))).toBe(
      "transitional"
    );
  });

  it("names neither class for a root of another vocabulary", () => {
    expect(
      conformanceOf(rootOf('<w:document xmlns:w="urn:other"/>'))
    ).toBeNull();
  });
});

describe("bindsWritingPrefix", () => {
  it("sees whether the root binds w to the transitional namespace", () => {
    expect(bindsWritingPrefix(rootOf(`<w:document xmlns:w="${W_NS}"/>`))).toBe(
      true
    );
    expect(
      bindsWritingPrefix(rootOf(`<w:document xmlns:w="${STRICT_W_NS}"/>`))
    ).toBe(false);
  });

  it("says no to a root that binds the namespace to another prefix", () => {
    expect(bindsWritingPrefix(rootOf(`<p:document xmlns:p="${W_NS}"/>`))).toBe(
      false
    );
  });

  it("says no to a root that binds the namespace as the default one", () => {
    expect(bindsWritingPrefix(rootOf(`<document xmlns="${W_NS}"/>`))).toBe(
      false
    );
  });

  it("says yes where another prefix stands beside the one we write", () => {
    expect(
      bindsWritingPrefix(
        rootOf(`<p:document xmlns:p="${W_NS}" xmlns:w="${W_NS}"/>`)
      )
    ).toBe(true);
  });
});
