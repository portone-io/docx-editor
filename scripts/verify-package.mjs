import { execFile } from "node:child_process";
import {
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

const consumerApp = `import "@portone-io/docx-editor/styles.css";
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
} from "@portone-io/docx-editor";
import {
  canSetLineSpacing,
  increaseListLevel,
  insertTable,
  toggleBold,
} from "@portone-io/docx-editor/commands";
import { exportDocx, importDocx } from "@portone-io/docx-editor/core";
import { addRowAfter, canSetCellBorderColor } from "@portone-io/docx-editor/table";
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
import { DEFAULT_COLORS, DocxEditor, docxSchema, downloadDocx } from "@portone-io/docx-editor";
import { increaseListLevel, toggleBold } from "@portone-io/docx-editor/commands";
import { exportDocx, importDocx } from "@portone-io/docx-editor/core";
import { addRowAfter, canSetCellBorderColor } from "@portone-io/docx-editor/table";

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

async function verify() {
  const workDir = await mkdtemp(join(tmpdir(), "docx-editor-verify-"));
  const consumerDir = join(workDir, "consumer");
  process.stdout.write(`verifying the package in ${workDir}\n`);

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
    ].map((name) => {
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
            "@portone-io/docx-editor": `file:${tarball}`,
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
