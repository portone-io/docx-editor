import { existsSync } from "node:fs";
import { cp, readdir, readFile, rm } from "node:fs/promises";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

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

const SOURCE_SUFFIXES = [".ts", ".tsx", "/index.ts", "/index.tsx"];

/** The path an emitted file uses to import another one, or null when nothing answers to it */
function emittedPath(importer, specifier) {
  const base = resolve(dirname(importer), specifier);
  const source = SOURCE_SUFFIXES.map((suffix) => `${base}${suffix}`).find(
    (candidate) => existsSync(candidate)
  );
  if (source === undefined) return null;
  const path = relative(dirname(importer), source).replace(/\.tsx?$/, ".js");
  return path.startsWith(".") ? path : `./${path}`;
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
    builder.onResolve({ filter: /^\.\.?\// }, (args) => {
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

await cp(
  join(packageRoot, "src/styles/editor.css"),
  join(distDir, "styles.css")
);
