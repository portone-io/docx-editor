import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execute = promisify(execFile);
const repository = fileURLToPath(new URL("..", import.meta.url));
const library = "@portone/docx-editor";

test("Changesets bumps the library while preserving the already published demo pins", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "site-version-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  for (const path of [
    "package.json",
    "pnpm-workspace.yaml",
    "site/package.json",
    "demo/package.json",
  ]) {
    await mkdir(join(root, path, ".."), { recursive: true });
    await copyFile(join(repository, path), join(root, path));
  }
  const read = async (path) =>
    JSON.parse(await readFile(join(root, path), "utf8"));
  const before = await read("package.json");
  const pins = await Promise.all(
    ["site", "demo"].map(
      async (dir) => (await read(`${dir}/package.json`)).dependencies[library]
    )
  );
  const config = JSON.parse(
    await readFile(join(repository, ".changeset/config.json"), "utf8")
  );
  // Changelog generation is unrelated to dependency versioning and can contact GitHub.
  config.changelog = false;
  await mkdir(join(root, ".changeset"));
  await writeFile(join(root, ".changeset/config.json"), JSON.stringify(config));
  await writeFile(
    join(root, ".changeset/release.md"),
    `---\n"${library}": minor\n---\nRelease preparation probe.\n`
  );
  await execute(
    process.execPath,
    [join(repository, "node_modules/@changesets/cli/bin.js"), "version"],
    { cwd: root }
  );
  assert.notEqual((await read("package.json")).version, before.version);
  for (const [i, dir] of ["site", "demo"].entries()) {
    assert.equal(
      (await read(`${dir}/package.json`)).dependencies[library],
      pins[i]
    );
  }
});
