// @vitest-environment node
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const srcDir = dirname(fileURLToPath(import.meta.url));
const packageDir = join(srcDir, "..");

/**
 * A folder may import itself and strictly lower ranks only. page and table
 * share a rank, so they may not import each other. The subfolders inside
 * editor/ are navigation only and count as editor.
 */
const FOLDER_RANKS: Readonly<Record<string, number>> = {
  model: 0,
  styles: 1,
  ooxml: 2,
  numbering: 3,
  schema: 4,
  docx: 5,
  page: 6,
  table: 6,
  editor: 7,
  ui: 8,
  "(root)": 9,
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function productionFiles(dir: string): string[] {
  const files: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "__testing__") continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...productionFiles(path));
    else if (/\.tsx?$/.test(entry.name) && !entry.name.includes(".test.")) {
      files.push(path);
    }
  }
  return files;
}

function folderOf(file: string): string {
  const rel = relative(srcDir, file);
  return rel.includes(sep) ? rel.split(sep)[0] : "(root)";
}

function resolveImport(file: string, specifier: string): string {
  const base = resolve(dirname(file), specifier);
  const candidates = [`${base}.ts`, `${base}.tsx`, join(base, "index.ts")];
  const found = candidates.find((candidate) => existsSync(candidate));
  if (!found) {
    throw new Error(
      `${relative(srcDir, file)} imports "${specifier}", which resolves to no file`
    );
  }
  return found;
}

function importsOf(file: string): { specifier: string; target: string }[] {
  return [...readFileSync(file, "utf8").matchAll(/from "(\.[^"]+)"/g)].map(
    ([, specifier]) => ({
      specifier,
      target: resolveImport(file, specifier),
    })
  );
}

function entryFiles(): string[] {
  const pkg: unknown = JSON.parse(
    readFileSync(join(packageDir, "package.json"), "utf8")
  );
  if (!isRecord(pkg) || !isRecord(pkg.exports)) {
    throw new Error("package.json declares no exports");
  }
  return Object.values(pkg.exports)
    .filter(
      (target): target is string =>
        typeof target === "string" && /\.tsx?$/.test(target)
    )
    .map((target) => resolve(packageDir, target));
}

const files = productionFiles(srcDir);
const fileSet = new Set(files);

describe("the folder layering", () => {
  it("scans the production tree", () => {
    expect(files.length).toBeGreaterThanOrEqual(90);
    expect(files).toContain(join(srcDir, "DocxEditor.tsx"));
  });

  it("crosses folder boundaries downward only", () => {
    const violations: string[] = [];
    let crossEdges = 0;
    for (const file of files) {
      const fromFolder = folderOf(file);
      const fromRank = FOLDER_RANKS[fromFolder];
      expect(
        fromRank,
        `"${fromFolder}" (${relative(srcDir, file)}) has no rank; place new folders deliberately`
      ).toBeDefined();
      for (const { specifier, target } of importsOf(file)) {
        const toFolder = folderOf(target);
        if (toFolder === fromFolder) continue;
        crossEdges += 1;
        const toRank = FOLDER_RANKS[toFolder];
        expect(
          toRank,
          `"${toFolder}" (imported by ${relative(srcDir, file)}) has no rank; place new folders deliberately`
        ).toBeDefined();
        if (!(toRank < fromRank)) {
          violations.push(
            `${relative(srcDir, file)} (${fromFolder}, rank ${fromRank}) imports "${specifier}" (${toFolder}, rank ${toRank})`
          );
        }
      }
    }
    expect(crossEdges).toBeGreaterThan(0);
    expect(violations, violations.join("\n")).toEqual([]);
  });

  it("reaches every production file from the entry points", () => {
    const entries = entryFiles();
    expect(entries.length).toBeGreaterThan(0);

    const reached = new Set<string>();
    const visit = (file: string): void => {
      if (reached.has(file)) return;
      reached.add(file);
      for (const { target } of importsOf(file)) {
        if (fileSet.has(target)) visit(target);
      }
    };
    for (const entry of entries) visit(entry);

    const orphans = files
      .filter((file) => !reached.has(file))
      .map((file) => relative(srcDir, file));
    expect(
      orphans,
      `unreachable from every entry:\n${orphans.join("\n")}`
    ).toEqual([]);
  });
});
