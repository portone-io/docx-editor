import { spawnSync } from "node:child_process";
import { cp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { dirname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import {
  emittedPath,
  RELATIVE_SPECIFIER,
  rewriteDeclarationSpecifiers,
} from "./module-specifiers.mjs";

const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const srcDir = join(packageRoot, "src");
const distDir = join(packageRoot, "dist");

const manifest = JSON.parse(
  await readFile(join(packageRoot, "package.json"), "utf8")
);

/** Everything a consumer installs itself stays an import rather than being copied in */
const external = [
  ...Object.keys(manifest.dependencies ?? {}),
  ...Object.keys(manifest.peerDependencies ?? {}),
].flatMap((name) => [name, `${name}/*`]);

/** Every module the package ships, which is all of `src` bar the tests and their helpers */
async function moduleFiles(dir) {
  const files = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name !== "__testing__")
        files.push(...(await moduleFiles(path)));
    } else if (/\.tsx?$/.test(entry.name) && !entry.name.includes(".test.")) {
      files.push(path);
    }
  }
  return files;
}

async function declarationFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true, recursive: true });
  return entries
    .filter((entry) => entry.isFile() && entry.name.endsWith(".d.ts"))
    .map((entry) => join(entry.parentPath, entry.name));
}

/**
 * Leaves the module graph as it is instead of bundling it, by handing every import inside
 * the package back as an external one.
 *
 * Bundling flattens modules into a shared chunk, where one module's impure top level
 * statement holds every other module in the chunk in place: a consumer importing a single
 * helper had tens of kilobytes it never called pulled in behind it. Emitted one to one, the
 * consumer's own bundler decides what it reaches.
 *
 * The extension is added here because esbuild writes an import path out as it found it and
 * TypeScript sources leave it off, which Node would not resolve.
 */
const keepModules = {
  name: "keep-modules",
  setup(builder) {
    builder.onResolve({ filter: RELATIVE_SPECIFIER }, (args) => {
      if (args.kind === "entry-point") return null;
      const path = emittedPath(args.importer, args.path);
      if (path === null) {
        return {
          errors: [
            {
              text: `"${args.path}" imported from ${relative(srcDir, args.importer)} resolves to no source file`,
            },
          ],
        };
      }
      return { path, external: true };
    });
  },
};

await rm(distDir, { recursive: true, force: true });

await build({
  absWorkingDir: packageRoot,
  entryPoints: await moduleFiles(srcDir),
  outdir: "dist",
  outbase: "src",
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  jsx: "automatic",
  external,
  plugins: [keepModules],
  logLevel: "warning",
});

const tsc = join(
  dirname(createRequire(import.meta.url).resolve("typescript/package.json")),
  "bin/tsc"
);
const declarations = spawnSync(
  process.execPath,
  [tsc, "-p", "tsconfig.build.json"],
  { cwd: packageRoot, stdio: "inherit" }
);
if (declarations.status !== 0)
  throw new Error("tsc could not emit the declarations");

/**
 * Declarations name each other the way the sources do, which TypeScript's `node16` and
 * `nodenext` resolution reject, so they are given the same paths the `.js` output imports through.
 */
const rewrites = [];
const problems = [];
for (const file of await declarationFiles(distDir)) {
  const importer = join(srcDir, relative(distDir, file));
  try {
    const text = rewriteDeclarationSpecifiers(
      await readFile(file, "utf8"),
      (specifier) => emittedPath(importer, specifier)
    );
    rewrites.push({ file, text });
  } catch (error) {
    problems.push(`${relative(distDir, file)}: ${error.message}`);
  }
}
if (problems.length > 0)
  throw new Error(
    `the declarations could not be given emitted paths:\n${problems.join("\n")}`
  );
for (const { file, text } of rewrites) await writeFile(file, text);

await cp(
  join(packageRoot, "src/styles/editor.css"),
  join(distDir, "styles.css")
);
