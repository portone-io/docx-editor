// @vitest-environment node
import { execFile } from "node:child_process";
import { mkdtemp, readdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const run = promisify(execFile);
const packageDir = join(dirname(fileURLToPath(import.meta.url)), "..");

/**
 * Every path the published `exports` map resolves to, plus the files a consumer
 * reads before installing. The documentation itself is not among them: the site
 * serves it, and the README links there.
 */
const requiredEntries = [
  "package/package.json",
  "package/LICENSE",
  "package/README.md",
  "package/CHANGELOG.md",
  "package/CONTRIBUTING.md",
  "package/assets/editor.png",
  "package/dist/styles.css",
  "package/dist/index.js",
  "package/dist/index.d.ts",
  "package/dist/core.js",
  "package/dist/core.d.ts",
  "package/dist/editor/commands/index.js",
  "package/dist/editor/commands/index.d.ts",
  "package/dist/table/index.js",
  "package/dist/table/index.d.ts",
];

/**
 * The sources are left behind on purpose: nothing in the published `exports`
 * map points at them and no source map asks for them, so they would only make
 * the tarball larger and offer a second, unsupported way in.
 */
const forbiddenEntries = [
  { what: "the sources", pattern: /^package\/src\// },
  { what: "the docx fixtures", pattern: /__fixtures__/ },
  { what: "the unit tests", pattern: /\.test\./ },
  // A declaration emitted for a test helper would put `__testing__` in the public type surface
  { what: "the test helpers", pattern: /__testing__/ },
  { what: "the browser specs", pattern: /^package\/e2e\// },
  { what: "the development demo", pattern: /^package\/demo\// },
  { what: "the documentation site", pattern: /^package\/site\// },
  // The user-facing documentation is the site's pages; `docs/` holds contributor guides only
  { what: "the contributor guides", pattern: /^package\/docs\// },
  { what: "the OOXML schemas", pattern: /^package\/spec\// },
  { what: "the internal notes", pattern: /\/(PROVENANCE|PLAN)\.md$/ },
  {
    what: "the test configuration",
    pattern: /(vitest|playwright)\.config\./,
  },
];

const DIST = "package/dist/";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

function exportSubpaths(exportsField: unknown): string[] {
  return isRecord(exportsField) ? Object.keys(exportsField) : [];
}

function exportTargets(exportsField: unknown): string[] {
  if (!isRecord(exportsField)) return [];
  return Object.values(exportsField).flatMap((subpath) => {
    if (typeof subpath === "string") return [subpath];
    if (!isRecord(subpath)) return [];
    return Object.values(subpath).flatMap((condition) =>
      typeof condition === "string" ? [condition] : []
    );
  });
}

function dependencyRanges(manifest: Record<string, unknown>): string[] {
  return ["dependencies", "peerDependencies", "devDependencies"].flatMap(
    (field) => {
      const ranges = manifest[field];
      if (!isRecord(ranges)) return [];
      return Object.entries(ranges).map(
        ([name, range]) => `${name}@${String(range)}`
      );
    }
  );
}

let workDir = "";
let entries: string[] = [];
let manifest: Record<string, unknown> = {};
let developmentSubpaths: string[] = [];

async function readManifest(read: () => Promise<string>) {
  const parsed: unknown = JSON.parse(await read());
  if (!isRecord(parsed)) throw new Error("the manifest is not an object");
  return parsed;
}

beforeAll(async () => {
  workDir = await mkdtemp(join(tmpdir(), "docx-editor-pack-"));
  /** Packing has to produce the output itself, the way a release off a clean checkout would */
  await rm(join(packageDir, "dist"), { recursive: true, force: true });
  await run("pnpm", ["pack", "--pack-destination", workDir], {
    cwd: packageDir,
  });

  const packed = (await readdir(workDir)).filter((name) =>
    name.endsWith(".tgz")
  );
  expect(packed).toHaveLength(1);
  const tarball = join(workDir, packed[0]);

  const listing = await run("tar", ["-tzf", tarball]);
  entries = listing.stdout.split("\n").filter((line) => line.length > 0);

  manifest = await readManifest(async () => {
    const extracted = await run("tar", [
      "-xzOf",
      tarball,
      "package/package.json",
    ]);
    return extracted.stdout;
  });
  const source = await readManifest(() =>
    readFile(join(packageDir, "package.json"), "utf8")
  );
  developmentSubpaths = exportSubpaths(source.exports);
}, 180_000);

afterAll(async () => {
  if (workDir) await rm(workDir, { recursive: true, force: true });
});

describe("the published tarball", () => {
  it("holds files at all", () => {
    expect(entries.length).toBeGreaterThan(0);
    expect(entries).toContain("package/package.json");
  });

  it("ships every file the exports map resolves to", () => {
    for (const entry of requiredEntries) {
      expect(entries).toContain(entry);
    }
  });

  it.each(forbiddenEntries)("leaves $what behind", ({ pattern }) => {
    expect(entries.filter((entry) => pattern.test(entry))).toEqual([]);
  });

  /**
   * The build emits the module graph one to one rather than bundling it, so what a
   * consumer's bundler is handed is the modules the declarations describe
   */
  it("ships the module behind every declaration, and a declaration for every module", () => {
    const distPaths = (suffix: string) =>
      entries
        .filter((entry) => entry.startsWith(DIST) && entry.endsWith(suffix))
        .map((entry) => entry.slice(DIST.length, -suffix.length))
        .sort();
    const modules = distPaths(".js");
    expect(modules.length).toBeGreaterThan(4);
    expect(distPaths(".d.ts")).toEqual(modules);
  });

  it("offers the subpaths the development exports map declares", () => {
    expect(developmentSubpaths.length).toBeGreaterThan(0);
    expect(exportSubpaths(manifest.exports)).toEqual(developmentSubpaths);
  });

  it("points every export at the built output", () => {
    const targets = exportTargets(manifest.exports);
    expect(targets.length).toBeGreaterThan(0);
    for (const target of targets) {
      expect(target).toMatch(/^\.\/dist\//);
    }
  });

  it("resolves the workspace-only dependency protocols", () => {
    const ranges = dependencyRanges(manifest);
    expect(ranges.length).toBeGreaterThan(0);
    for (const range of ranges) {
      expect(range).not.toMatch(/catalog:|workspace:/);
    }
  });
});
