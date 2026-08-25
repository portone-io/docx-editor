import { execFile } from "node:child_process";
import {
  copyFile,
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const packageRoot = fileURLToPath(new URL("..", import.meta.url));

/**
 * The React the consumer installs, overriding the range the manifest declares.
 * The package supports more than one React major, and a single verification run can only
 * install one of them, so CI sets this to check each supported major in turn.
 */
const REACT_RANGE = process.env.DOCX_EDITOR_REACT_RANGE;

/** What `REACT_RANGE` stands in for. `@types/react-dom` is not among them: no consumer file needs it */
const REACT_PACKAGES = new Set(["react", "react-dom", "@types/react"]);

/** The document the render smoke opens, small enough to carry into the consumer project */
const RENDER_FIXTURE = "demo.docx";

async function step(what, action) {
  process.stdout.write(`  ${what}\n`);
  try {
    return await action();
  } catch (cause) {
    const detail = cause.stderr || cause.stdout || cause.message || cause;
    throw new Error(`${what}\n\n${String(detail).trim()}`);
  }
}

const consumerTsconfig = {
  compilerOptions: {
    target: "es2022",
    lib: ["dom", "dom.iterable", "esnext"],
    module: "preserve",
    moduleResolution: "bundler",
    jsx: "react-jsx",
    strict: true,
    noEmit: true,
    skipLibCheck: false,
    types: [],
  },
  include: ["app.tsx", "env.d.ts"],
};

const consumerApp = `import "@portone/docx-editor/styles.css";
import {
  type ColorRow,
  DEFAULT_COLORS,
  DEFAULT_FONTS,
  DocxEditor,
  type DocxEditorHandle,
  type DocxSource,
  docxSchema,
  downloadDocx,
  type DownloadDocxResult,
} from "@portone/docx-editor";
import {
  canSetLineSpacing,
  increaseListLevel,
  insertTable,
  toggleBold,
} from "@portone/docx-editor/commands";
import { exportDocx, importDocx } from "@portone/docx-editor/core";
import { addRowAfter, canSetCellBorderColor } from "@portone/docx-editor/table";
import { useRef } from "react";

export function Consumer({ document }: { document: DocxSource }) {
  const editor = useRef<DocxEditorHandle | null>(null);
  return (
    <DocxEditor
      ref={editor}
      document={document}
      mode={{ kind: "edit", locking: true }}
      presets={{ fonts: DEFAULT_FONTS, colors: DEFAULT_COLORS }}
    />
  );
}

export function roundTrip(bytes: ArrayBuffer): Uint8Array {
  const { doc, session } = importDocx(bytes);
  return exportDocx(doc, session);
}

export function save(editor: DocxEditorHandle | null): DownloadDocxResult {
  return downloadDocx(editor, { fileName: "contract" });
}

export const firstColorRow: ColorRow | undefined = DEFAULT_COLORS[0];
export const editing = [
  toggleBold,
  addRowAfter,
  increaseListLevel,
  insertTable({ rows: 2, columns: 3 }),
];
export const queries = [canSetLineSpacing, canSetCellBorderColor];
export const nodeNames = Object.keys(docxSchema.nodes);
`;

const consumerTypes = `declare module "*.css";
`;

const consumerSmoke = `import assert from "node:assert/strict";
import { DEFAULT_COLORS, DocxEditor, docxSchema, downloadDocx } from "@portone/docx-editor";
import { increaseListLevel, toggleBold } from "@portone/docx-editor/commands";
import { exportDocx, importDocx } from "@portone/docx-editor/core";
import { addRowAfter, canSetCellBorderColor } from "@portone/docx-editor/table";

const surface = {
  DEFAULT_COLORS,
  DocxEditor,
  docxSchema,
  downloadDocx,
  increaseListLevel,
  toggleBold,
  exportDocx,
  importDocx,
  addRowAfter,
  canSetCellBorderColor,
};

for (const [name, value] of Object.entries(surface)) {
  assert.ok(value, name + " is missing from the published entries");
}
`;

const consumerRender = `import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { JSDOM } from "jsdom";

const dom = new JSDOM("<!doctype html><html><body></body></html>", {
  // An origin of its own, without which the window refuses to hand out storage
  url: "https://consumer.test/",
  pretendToBeVisual: true,
});

// The package is imported only once the DOM globals are in place, the way a browser has them.
// Only what node has no answer for is taken over, which leaves the language builtins, the
// timers, and the clock as node's. Each one is read off the window rather than copied as a
// descriptor, because several are accessors that only answer to the window itself.
globalThis.window = dom.window;
for (const key of Object.getOwnPropertyNames(dom.window)) {
  if (key in globalThis) continue;
  globalThis[key] = dom.window[key];
}
globalThis.IS_REACT_ACT_ENVIRONMENT = true;

/** The download needs an object URL that comes back out again, which jsdom does not hand out */
const objectURLs = [];
globalThis.URL.createObjectURL = (blob) => {
  const url = "blob:consumer/" + objectURLs.length + "#" + blob.size;
  objectURLs.push(url);
  return url;
};
globalThis.URL.revokeObjectURL = () => {};

/** The download holds its object URL for a minute before revoking it; that timer must not hold the run open */
const schedule = globalThis.setTimeout;
globalThis.setTimeout = (handler, delay, ...rest) => {
  const timer = schedule(handler, delay, ...rest);
  timer.unref();
  return timer;
};

const { act, createElement } = await import("react");
const { createRoot } = await import("react-dom/client");
const { DocxEditor, downloadDocx } = await import("@portone/docx-editor");

const bytes = new Uint8Array(await readFile("${RENDER_FIXTURE}"));
const host = document.createElement("div");
document.body.append(host);

const editor = { current: null };
const root = createRoot(host);
await act(async () => {
  root.render(createElement(DocxEditor, { document: bytes, ref: editor }));
});

const handle = editor.current;
assert.ok(handle, "the editor handed the consumer's ref no handle");
assert.equal(
  typeof handle.exportBytes,
  "function",
  "the handle exposes no exportBytes"
);
assert.ok(
  handle.exportBytes().length > 0,
  "the mounted editor exported an empty document"
);

const download = downloadDocx(handle, { fileName: "contract" });
assert.notEqual(
  download.status,
  "unavailable",
  "downloadDocx could not reach the mounted editor"
);
assert.deepEqual(download, {
  status: "exported",
  fileName: "contract.docx",
  byteLength: handle.exportBytes().length,
});
assert.equal(objectURLs.length, 1, "the download made no file");

await act(async () => root.unmount());
dom.window.close();
`;

async function verify() {
  const workDir = await mkdtemp(join(tmpdir(), "docx-editor-verify-"));
  const consumerDir = join(workDir, "consumer");
  process.stdout.write(
    `verifying the package in ${workDir}\n` +
      (REACT_RANGE ? `  against react@${REACT_RANGE}\n` : "")
  );

  await step("building the package", () =>
    run("pnpm", ["build"], { cwd: packageRoot })
  );

  const tarball = await step("packing the package", async () => {
    await run("pnpm", ["pack", "--pack-destination", workDir], {
      cwd: packageRoot,
    });
    const packed = (await readdir(workDir)).find((name) =>
      name.endsWith(".tgz")
    );
    if (!packed) throw new Error("pnpm pack left no tarball behind");
    return join(workDir, packed);
  });

  const manifest = await step("reading the packed manifest", async () => {
    const { stdout } = await run("tar", [
      "-xzOf",
      tarball,
      "package/package.json",
    ]);
    return JSON.parse(stdout);
  });

  /** The peers a consumer has to install itself, using the published ranges. */
  const declared = {
    ...manifest.peerDependencies,
    ...manifest.devDependencies,
  };
  const consumerDependencies = Object.fromEntries(
    [
      ...Object.keys(manifest.peerDependencies ?? {}),
      "react-dom",
      "@types/react",
      "typescript",
      "esbuild",
      "jsdom",
    ].map((name) => {
      if (REACT_RANGE && REACT_PACKAGES.has(name)) return [name, REACT_RANGE];
      const range = declared[name];
      if (!range) throw new Error(`the packed manifest names no ${name}`);
      return [name, range];
    })
  );

  await step("writing a consumer project", async () => {
    await mkdir(consumerDir, { recursive: true });
    await writeFile(
      join(consumerDir, "package.json"),
      `${JSON.stringify(
        {
          name: "docx-editor-consumer",
          version: "0.0.0",
          private: true,
          type: "module",
          dependencies: {
            "@portone/docx-editor": `file:${tarball}`,
            ...consumerDependencies,
          },
        },
        null,
        2
      )}\n`
    );
    await writeFile(
      join(consumerDir, "tsconfig.json"),
      `${JSON.stringify(consumerTsconfig, null, 2)}\n`
    );
    /** esbuild unpacks its binary in an install script, which pnpm runs only where it is allowed */
    await writeFile(
      join(consumerDir, "pnpm-workspace.yaml"),
      "allowBuilds:\n  esbuild: true\n"
    );
    await writeFile(join(consumerDir, "app.tsx"), consumerApp);
    await writeFile(join(consumerDir, "env.d.ts"), consumerTypes);
    await writeFile(join(consumerDir, "smoke.mjs"), consumerSmoke);
    await writeFile(join(consumerDir, "render.mjs"), consumerRender);
    await copyFile(
      join(packageRoot, "__fixtures__", RENDER_FIXTURE),
      join(consumerDir, RENDER_FIXTURE)
    );
  });

  await step("installing the tarball and its peers", () =>
    run("pnpm", ["install"], { cwd: consumerDir })
  );

  await step("typechecking against the shipped declarations", () =>
    run(join(consumerDir, "node_modules/.bin/tsc"), ["-p", "tsconfig.json"], {
      cwd: consumerDir,
    })
  );

  await step("bundling the consumer", () =>
    run(
      join(consumerDir, "node_modules/.bin/esbuild"),
      [
        "app.tsx",
        "--bundle",
        "--format=esm",
        "--outdir=out",
        '--define:process.env.NODE_ENV="production"',
      ],
      { cwd: consumerDir }
    )
  );

  await step("loading the entries in node", () =>
    run(process.execPath, ["smoke.mjs"], { cwd: consumerDir })
  );

  await step("rendering the editor over a fixture document", () =>
    run(process.execPath, ["render.mjs"], { cwd: consumerDir })
  );

  const sizes = await step("checking what the bundle came to", async () => {
    const script = await stat(join(consumerDir, "out/app.js"));
    const stylesheet = await readFile(join(consumerDir, "out/app.css"), "utf8");
    if (script.size === 0) throw new Error("the bundled script is empty");
    if (!stylesheet.includes(".docx-editor-"))
      throw new Error(
        "the bundled stylesheet holds none of the editor's rules"
      );
    return { script: script.size, stylesheet: stylesheet.length };
  });

  const tarballSize = (await stat(tarball)).size;
  process.stdout.write(
    `\nthe package installs and builds outside the repository\n` +
      `  tarball    ${Math.round(tarballSize / 1024)} kB\n` +
      `  bundle     ${Math.round(sizes.script / 1024)} kB of script, ` +
      `${Math.round(sizes.stylesheet / 1024)} kB of styles\n`
  );

  await rm(workDir, { recursive: true, force: true });
}

try {
  await verify();
} catch (error) {
  process.stderr.write(`\nverify:package failed while ${error.message}\n`);
  process.exitCode = 1;
}
