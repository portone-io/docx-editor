// @vitest-environment node
import { execFile } from "node:child_process";
import { readFileSync } from "node:fs";
import { rm } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { Extractor, ExtractorConfig } from "@microsoft/api-extractor";
import { beforeAll, describe, expect, it } from "vitest";

const run = promisify(execFile);
const packageDir = join(dirname(fileURLToPath(import.meta.url)), "..");

/** The api-extractor configuration that reports on each published subpath */
const configNames: Readonly<Record<string, string>> = {
  ".": "index",
  "./core": "core",
  "./commands": "commands",
  "./table": "table",
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Every published subpath that resolves to declarations, and the file it names */
function declaredEntries(): Record<string, string> {
  const manifest: unknown = JSON.parse(
    readFileSync(join(packageDir, "package.json"), "utf8")
  );
  const published = isRecord(manifest) ? manifest.publishConfig : undefined;
  const subpaths = isRecord(published) ? published.exports : undefined;
  if (!isRecord(subpaths))
    throw new Error("publishConfig.exports must hold an object");

  return Object.fromEntries(
    Object.entries(subpaths).flatMap(([subpath, conditions]) => {
      const types = isRecord(conditions) ? conditions.types : undefined;
      return typeof types === "string" ? [[subpath, types]] : [];
    })
  );
}

const entries = declaredEntries();

beforeAll(async () => {
  // The declarations have to be the build's own, the way a release off a clean checkout would be
  await rm(join(packageDir, "dist"), { recursive: true, force: true });
  await run("pnpm", ["build"], { cwd: packageDir });
}, 180_000);

describe("the published type surface", () => {
  it("is reported on for every subpath that ships declarations", () => {
    expect(Object.keys(configNames).sort()).toEqual(
      Object.keys(entries).sort()
    );
  });

  it.each(Object.entries(entries))(
    "the declaration report of %s is the committed one",
    (subpath, types) => {
      const name = configNames[subpath];
      if (name === undefined)
        throw new Error(`${subpath} has no api-extractor configuration`);

      const config = ExtractorConfig.loadFileAndPrepare(
        join(packageDir, "api-extractor", `${name}.json`)
      );
      expect(config.mainEntryPointFilePath).toBe(join(packageDir, types));

      const result = Extractor.invoke(config, { localBuild: false });
      expect(result.succeeded).toBe(true);
      expect(result.apiReportChanged).toBe(false);
    }
  );
});
