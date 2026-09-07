// @vitest-environment node
import { execFile } from "node:child_process";
import { readdirSync } from "node:fs";
import { readFile, rm } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { JSDOM } from "jsdom";
import { beforeAll, describe, expect, it } from "vitest";

const run = promisify(execFile);
const packageDir = join(dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = join(packageDir, "__fixtures__");

const fixtureNames = readdirSync(fixturesDir).filter((name) =>
  name.endsWith(".docx")
);

/** The parser a server hands in, which is the only one anything here has to read with */
const { DOMParser } = new JSDOM().window;
const xmlParser = new DOMParser();

/**
 * The published core entry, opened where a server opens it.
 *
 * The rest of the suite runs under jsdom or installs the globals it wants, so this is the one
 * place a runtime holding no DOM at all is what the entry meets. It is the built output rather
 * than the source, because what a consumer installs is what has to hold.
 */
type CoreEntry = typeof import("../src/core");

let core: CoreEntry;

beforeAll(async () => {
  // The output has to be the build's own, the way a release off a clean checkout would be.
  // Every spec in this folder writes `dist`, which is why `test:package` runs one file at a time
  await rm(join(packageDir, "dist"), { recursive: true, force: true });
  await run("pnpm", ["build"], { cwd: packageDir });
  core = await import(pathToFileURL(join(packageDir, "dist/core.js")).href);
}, 180_000);

/** The code a refusal carries, the way a consumer switching on it would read it */
function refusalCode(work: () => unknown): string {
  try {
    work();
  } catch (error) {
    if (error instanceof core.DocxImportError) return error.code;
    throw error;
  }
  throw new Error("the read was expected to be refused");
}

describe("the core entry in a runtime with no DOM", () => {
  it("runs with neither of the globals a browser would have lent it", () => {
    expect("DOMParser" in globalThis).toBe(false);
    expect("Node" in globalThis).toBe(false);
  });

  // Reaching for a global that is not there throws a `ReferenceError` naming it, which a
  // consumer cannot tell apart from a file that arrived damaged
  it("refuses to open a file without a parser with a stable code", async () => {
    const bytes = await readFile(join(fixturesDir, "demo.docx"));

    expect(refusalCode(() => core.importDocx(bytes))).toBe("no-xml-parser");
  });

  it.each(fixtureNames)(
    "opens and writes back %s with a jsdom parser passed as an option",
    async (name) => {
      const bytes = await readFile(join(fixturesDir, name));

      const { doc, session } = core.importDocx(bytes, { xmlParser });

      expect(doc.childCount).toBeGreaterThan(0);
      expect(core.exportDocx(doc, session, { xmlParser })).not.toHaveLength(0);
    }
  );
});
