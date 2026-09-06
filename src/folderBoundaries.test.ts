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

/**
 * One import or re-export statement, as the clause it names and the file it reads.
 *
 * The clause is matched without a `;` in it so that a statement can never run into the next one,
 * which is what a lazy match over a file starting with several type-only imports would otherwise
 * do.
 */
const IMPORT_STATEMENT =
  /(?:^|\n)\s*(?:import|export)\s+([^;]*?)\s+from\s+"(\.[^"]+)";/g;

/**
 * The imports that survive into what runs, which are the ones a cycle can be built out of.
 *
 * A statement importing nothing but types is erased before anything runs and cannot hold a cycle
 * up, so it is left out. A statement mixing a type in with a value is kept whole: the module is
 * still fetched, and only writing it as `import type` would take it out of the graph.
 */
function runtimeImportsOf(file: string): string[] {
  return [...readFileSync(file, "utf8").matchAll(IMPORT_STATEMENT)]
    .filter(([, clause]) => !/^type\b/.test(clause.trim()))
    .map(([, , specifier]) => resolveImport(file, specifier));
}

/** The first cycle the runtime import graph holds, as the path around it. Empty when it holds none */
function firstRuntimeCycle(): string[] {
  const done = new Set<string>();
  const path: string[] = [];
  const onPath = new Set<string>();

  const walk = (file: string): string[] => {
    onPath.add(file);
    path.push(file);
    for (const target of runtimeImportsOf(file)) {
      if (!fileSet.has(target)) continue;
      if (onPath.has(target)) {
        return [...path.slice(path.indexOf(target)), target];
      }
      if (done.has(target)) continue;
      const found = walk(target);
      if (found.length > 0) return found;
    }
    path.pop();
    onPath.delete(file);
    done.add(file);
    return [];
  };

  for (const file of files) {
    if (done.has(file)) continue;
    const found = walk(file);
    if (found.length > 0) return found;
  }
  return [];
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

  /**
   * Two modules that read each other are evaluated in whichever order the first importer happens
   * to reach them, and whatever the module entered second reads at evaluation time is not there
   * yet. A list built out of what other modules export - `schema/guards` and its `EDIT_GUARDS` -
   * then holds a hole rather than raising anything, since the bundler lowers the declarations it
   * is built from to `var`.
   */
  it("holds no cycle among the imports that survive into what runs", () => {
    const cycle = firstRuntimeCycle().map((file) => relative(srcDir, file));

    expect(
      cycle,
      `these modules read each other:\n${cycle.join(" ->\n")}\nMove what both of them need into a module that imports neither, or make one of the two imports \`import type\`.`
    ).toEqual([]);
  });
});
