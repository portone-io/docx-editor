// @vitest-environment node
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import * as core from "./core";
import * as commands from "./editor/commands/index";
import * as root from "./index";
import * as table from "./table";

const packageDir = join(dirname(fileURLToPath(import.meta.url)), "..");

const MANIFEST_FILE = "api-manifest.json";

/**
 * Runtime exports only. A type-only export leaves nothing on the namespace object,
 * so nothing here guards one; api-extractor would be the tool for that.
 */
const ENTRY_POINTS: Readonly<Record<string, object>> = {
  ".": root,
  "./commands": commands,
  "./core": core,
  "./table": table,
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isStringArray(value: unknown): value is string[] {
  return (
    Array.isArray(value) && value.every((item) => typeof item === "string")
  );
}

function readJson(file: string): unknown {
  return JSON.parse(readFileSync(join(packageDir, file), "utf8"));
}

function readManifest(): Record<string, string[]> {
  const manifest = readJson(MANIFEST_FILE);
  if (!isRecord(manifest))
    throw new Error(`${MANIFEST_FILE} must hold an object keyed by subpath`);

  return Object.fromEntries(
    Object.entries(manifest).map(([subpath, names]) => {
      if (!isStringArray(names))
        throw new Error(`${MANIFEST_FILE}: ${subpath} must hold export names`);
      return [subpath, [...names].sort()];
    })
  );
}

function declaredJsSubpaths(): string[] {
  const manifest = readJson("package.json");
  const exports = isRecord(manifest) ? manifest.exports : undefined;
  if (!isRecord(exports)) throw new Error("package.json declares no exports");

  return Object.entries(exports)
    .filter(
      ([, target]) => typeof target === "string" && target.endsWith(".ts")
    )
    .map(([subpath]) => subpath)
    .sort();
}

const manifest = readManifest();

function diffOf(actual: string[], expected: string[]): string {
  return [
    ...actual
      .filter((name) => !expected.includes(name))
      .map((name) => `+ ${name}`),
    ...expected
      .filter((name) => !actual.includes(name))
      .map((name) => `- ${name}`),
  ].join("\n");
}

describe("the public API", () => {
  it("is pinned for every JavaScript entry point the package declares", () => {
    const declared = declaredJsSubpaths();
    expect(Object.keys(ENTRY_POINTS).sort()).toEqual(declared);
    expect(Object.keys(manifest).sort()).toEqual(declared);
  });

  describe.each(Object.entries(ENTRY_POINTS))("%s", (subpath, namespace) => {
    const actual = Object.keys(namespace).sort();
    const expected = manifest[subpath] ?? [];

    it("exports what the manifest pins", () => {
      // A pair of empty lists would match without pinning anything
      expect(actual.length).toBeGreaterThan(0);
      expect(expected.length).toBeGreaterThan(0);
      expect(
        actual,
        `The runtime exports of "${subpath}" no longer match ${MANIFEST_FILE}:\n${diffOf(actual, expected)}\nA change to the public API has to be deliberate and reviewed, so update ${MANIFEST_FILE} in the same commit as the change.`
      ).toEqual(expected);
    });
  });
});
