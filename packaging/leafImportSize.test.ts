// @vitest-environment node
import { execFile } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { build } from "esbuild";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const run = promisify(execFile);
const packageDir = join(dirname(fileURLToPath(import.meta.url)), "..");

const entries: Readonly<Record<string, string>> = {
  ".": "dist/index.js",
  "./commands": "dist/editor/commands/index.js",
  "./core": "dist/core.js",
  "./table": "dist/table/index.js",
};

function runtimeExports(): Record<string, string[]> {
  const value: unknown = JSON.parse(
    readFileSync(join(packageDir, "api-manifest.json"), "utf8")
  );
  if (typeof value !== "object" || value === null || Array.isArray(value))
    throw new Error("api-manifest.json must hold an object");

  return Object.fromEntries(
    Object.entries(value).map(([subpath, names]) => {
      if (
        !Array.isArray(names) ||
        !names.every((name): name is string => typeof name === "string")
      )
        throw new Error(`${subpath} must hold runtime export names`);
      return [subpath, names];
    })
  );
}

const exportsByEntry = runtimeExports();

/**
 * What a consumer importing one name off an entry is left holding.
 *
 * The build emits the module graph one to one, so an import reaches the module it names and
 * the ones that module itself reads. Bundling the entries instead flattens the modules into
 * shared chunks, where one impure top level statement holds every other module in the chunk
 * in place: `downloadDocx`, a dozen lines that put a blob in front of the user, came to 48 kB
 * that way, and `emuToPx`, one multiplication, to 22 kB.
 * The budgets are a few times what each costs today, so they catch that collapse rather than
 * a module gaining a line.
 *
 * The modules are the ones the bundle keeps any bytes of, because the same impure statement
 * pins a module at any size: `image.ts` reading the namespace table at its top level held
 * `names.js` and `xml.js` in place for a consumer that never wrote a drawing, and `emuToPx`
 * came to 817 bytes, well under its budget. A module in the received list that the name never
 * calls into is one such statement to take out, not a module to add. `index.js` stays in for
 * the ProseMirror classes it re-exports, because an import of an external package is kept as
 * it is written.
 */
const budgets = [
  {
    name: "downloadDocx",
    entry: "dist/index.js",
    bytes: 4096,
    modules: ["dist/download.js", "dist/index.js"],
  },
  {
    name: "emuToPx",
    entry: "dist/core.js",
    bytes: 1024,
    modules: ["dist/ooxml/image.js"],
  },
];

let workDir = "";

beforeAll(async () => {
  workDir = await mkdtemp(join(tmpdir(), "docx-editor-leaf-"));
  // The output has to be the build's own, the way a release off a clean checkout would be.
  // Every spec in this folder writes `dist`, which is why `test:package` runs one file at a time
  await rm(join(packageDir, "dist"), { recursive: true, force: true });
  await run("pnpm", ["build"], { cwd: packageDir });
}, 180_000);

afterAll(async () => {
  if (workDir) await rm(workDir, { recursive: true, force: true });
});

/**
 * What one imported name adds to a consumer's bundle, minified as one ships: the bytes, and
 * the modules under `dist` any of those bytes came from
 */
async function consumerBundle(
  name: string,
  entry: string
): Promise<{ bytes: number; modules: string[] }> {
  const file = join(workDir, `${name}.js`);
  await writeFile(
    file,
    `import { ${name} } from "${join(packageDir, entry)}";\n` +
      `globalThis.keep = ${name};\n`
  );
  const bundled = await build({
    absWorkingDir: packageDir,
    entryPoints: [file],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: "es2022",
    // What a consumer installs itself is not what is being measured
    packages: "external",
    minify: true,
    write: false,
    logLevel: "silent",
    metafile: true,
  });
  const [output] = Object.values(bundled.metafile.outputs);
  return {
    bytes: bundled.outputFiles[0].contents.length,
    modules: Object.keys(output.inputs)
      .filter((path) => path.startsWith("dist/"))
      .sort(),
  };
}

async function unusedImportSize(
  label: string,
  entry: string,
  names: readonly string[]
): Promise<number> {
  const file = join(workDir, `unused-${label}.js`);
  await writeFile(
    file,
    (names.length > 0
      ? `import { ${names.join(", ")} } from "${join(packageDir, entry)}";\n`
      : "") + "globalThis.keep = 1;\n"
  );
  const bundled = await build({
    entryPoints: [file],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: "es2022",
    packages: "external",
    minify: true,
    write: false,
    logLevel: "silent",
  });
  return bundled.outputFiles[0].contents.length;
}

describe("what one imported name costs a consumer", () => {
  it.each(budgets)(
    "$name off $entry stays under $bytes bytes",
    async ({ name, entry, bytes }) => {
      expect((await consumerBundle(name, entry)).bytes).toBeLessThan(bytes);
    }
  );

  it.each(budgets)(
    "$name off $entry keeps only the modules it reads",
    async ({ name, entry, modules }) => {
      expect((await consumerBundle(name, entry)).modules).toEqual(modules);
    }
  );

  it.each(
    Object.entries(entries).map(([subpath, entry], index) => ({
      subpath,
      entry,
      index,
    }))
  )(
    "drops every unused runtime export from $subpath",
    async ({ subpath, entry, index }) => {
      const names = exportsByEntry[subpath];
      if (names === undefined)
        throw new Error(`${subpath} is missing from api-manifest.json`);

      const baseline = await unusedImportSize(`baseline-${index}`, entry, []);
      const imported = await unusedImportSize(`entry-${index}`, entry, names);
      expect(imported).toBe(baseline);
    }
  );
});
