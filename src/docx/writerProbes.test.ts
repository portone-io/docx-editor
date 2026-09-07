// @vitest-environment jsdom
/**
 * Every command the package exports either writes something into an exported package or does
 * not, and this is where that question is answered for each of them.
 *
 * `WRITER_PROBES` runs the ones that write, before the export the schemas are read against
 * (`docx/exportSchemaValidation.test.ts`), and `NOT_A_WRITER` carries the rest with the reason
 * each reaches no writer. A command in neither list is one whose output nothing ever validates,
 * which is what the first test here fails on.
 *
 * The names are read off `api-manifest.json` rather than off the entry points, because that file
 * is where a change to the public surface has to be written down (`src/publicApi.test.ts`), so a
 * command cannot arrive without passing through it.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { NOT_A_WRITER, WRITER_PROBES } from "./__testing__/writerProbes";

const packageDir = join(dirname(fileURLToPath(import.meta.url)), "../..");

const MANIFEST_FILE = "api-manifest.json";

/** The entries whose commands the export battery is answerable for */
const ENTRY_POINTS = ["./commands", "./table"];

function isStringArray(value: unknown): value is string[] {
  return (
    Array.isArray(value) && value.every((item) => typeof item === "string")
  );
}

/** The runtime exports the manifest pins for those entries */
function pinnedExports(): string[] {
  const manifest: unknown = JSON.parse(
    readFileSync(join(packageDir, MANIFEST_FILE), "utf8")
  );
  if (typeof manifest !== "object" || manifest === null) {
    throw new Error(`${MANIFEST_FILE} must hold an object keyed by subpath`);
  }
  return ENTRY_POINTS.flatMap((subpath) => {
    const names = Reflect.get(manifest, subpath);
    if (!isStringArray(names)) {
      throw new Error(`${MANIFEST_FILE} pins no exports for "${subpath}"`);
    }
    return names;
  });
}

const exported = pinnedExports();
const probed = Object.keys(WRITER_PROBES);
const notWriters = Object.keys(NOT_A_WRITER);

describe("the writer probes", () => {
  it("answers for every export of `./commands` and `./table`", () => {
    // A pair of empty lists would agree with each other without answering anything
    expect(exported.length).toBeGreaterThan(0);
    expect(probed.length).toBeGreaterThan(0);

    const answered = new Set([...probed, ...notWriters]);
    const unanswered = exported.filter((name) => !answered.has(name));
    const stale = [...answered].filter((name) => !exported.includes(name));

    expect(
      unanswered,
      `exported but never run before an export is validated: ${unanswered.join(", ")}\nAdd a probe to WRITER_PROBES, or the name to NOT_A_WRITER with the reason it reaches no writer.`
    ).toEqual([]);
    expect(stale, `listed but no longer exported: ${stale.join(", ")}`).toEqual(
      []
    );
  });

  it("puts no name in both lists", () => {
    const both = probed.filter((name) => notWriters.includes(name));
    expect(
      both,
      `both probed and listed as reaching no writer: ${both.join(", ")}`
    ).toEqual([]);
  });

  it("runs at least one probe for every name it probes", () => {
    const empty = Object.entries(WRITER_PROBES)
      .filter(([, probes]) => probes.length === 0)
      .map(([name]) => name);

    expect(
      empty,
      `probed by an empty list, which runs nothing: ${empty.join(", ")}`
    ).toEqual([]);
  });
});
