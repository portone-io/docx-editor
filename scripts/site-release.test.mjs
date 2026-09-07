import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readFile,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { check, pin } from "./demo-library-pin.mjs";
import { commit, prepare, publishedVersion } from "./site-release.mjs";

const execute = promisify(execFile);
const repository = fileURLToPath(new URL("..", import.meta.url));
const library = "@portone/docx-editor";
const files = ["site/package.json", "demo/package.json", "pnpm-lock.yaml"];
const json = (root, path, value) =>
  writeFile(join(root, path), JSON.stringify(value));
const inputs = (root) =>
  Promise.all(files.map((path) => readFile(join(root, path), "utf8")));
const noWait = async () => {};

async function installed(root, version) {
  for (const dir of ["site", "demo"]) {
    await json(root, `${dir}/node_modules/${library}/package.json`, {
      name: library,
      version,
    });
  }
  await writeFile(join(root, "pnpm-lock.yaml"), `installed: ${version}\n`);
}
async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), "site-release-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await json(root, "package.json", { name: library, version: "0.3.0" });
  for (const dir of ["site", "demo"]) {
    await mkdir(join(root, dir, "node_modules", library), { recursive: true });
    await json(root, `${dir}/package.json`, {
      name: `${library}-${dir}`,
      version: "0.0.0",
      private: true,
      dependencies: { [library]: "0.2.1" },
    });
  }
  await installed(root, "0.2.1");
  return root;
}

test("the CLI accepts a newer root version but rejects source linked as the released demo", async (t) => {
  const root = await fixture(t);
  // Copy under scripts so the CLI finds this workspace, also exercising macOS /var aliases.
  await mkdir(join(root, "scripts"));
  await writeFile(
    join(root, "scripts/check.mjs"),
    await readFile(join(repository, "scripts/demo-library-pin.mjs"))
  );
  const { stdout } = await execute(process.execPath, [
    join(root, "scripts/check.mjs"),
  ]);
  assert.match(stdout, /@portone\/docx-editor@0\.2\.1 is installed/);
  await json(root, "package.json", { name: library, version: "0.2.1" });
  const entry = join(root, "demo/node_modules", library, "package.json");
  await rm(entry);
  await symlink(join(root, "package.json"), entry);
  await assert.rejects(check(root), /working tree/);
});

for (const version of ["latest", "0.3.0"]) {
  test(`preparing ${version} waits for the release and updates both consumers in CI`, async (t) => {
    const root = await fixture(t);
    let views = 0;
    let installs = 0;
    await pin(root, {
      version,
      wait: noWait,
      run: async (command, args) => {
        if (command === "npm") {
          assert.equal(args[1], `${library}@${version}`);
          if (++views === 1) throw new Error("version not visible yet");
          return { stdout: '"0.3.0"' };
        }
        assert.ok(args.includes("--no-frozen-lockfile"));
        assert.ok(args.includes("--ignore-scripts"));
        if (++installs === 1) throw new Error("tarball not visible yet");
        await installed(root, "0.3.0");
      },
    });
    assert.equal(await check(root), "0.3.0");
    assert.deepEqual([views, installs], [2, 2]);
  });
}

for (const failed of ["npm", "pnpm"]) {
  test(`${failed} failure stops after bounded retries and restores build inputs`, async (t) => {
    const root = await fixture(t);
    const before = await inputs(root);
    let failures = 0;
    await assert.rejects(
      pin(root, {
        attempts: 2,
        wait: noWait,
        run: async (command) => {
          if (command !== failed) return { stdout: '"0.3.0"' };
          failures++;
          if (command === "pnpm")
            await writeFile(join(root, "pnpm-lock.yaml"), "partial update");
          throw new Error("unavailable");
        },
      })
    );
    assert.equal(failures, 2);
    assert.deepEqual(await inputs(root), before);
  });
}

test("site preparation awaits installation and propagates install or build failures", async (t) => {
  const root = await fixture(t);
  for (const failed of [null, "install", "build"]) {
    const steps = [];
    const result = prepare(root, "0.3.0", {
      install: async (_root, options) => {
        assert.equal(options.version, "0.3.0");
        await new Promise((resolve) => setImmediate(resolve));
        steps.push("install");
        if (failed === "install") throw new Error(failed);
      },
      execute: async (_command, args) => {
        assert.deepEqual(steps, ["install"]);
        assert.deepEqual(args, ["build:site"]);
        steps.push("build");
        if (failed === "build") throw new Error(failed);
      },
    });
    if (failed) await assert.rejects(result, new RegExp(failed));
    else await result;
    assert.deepEqual(
      steps,
      failed === "install" ? ["install"] : ["install", "build"]
    );
  }
});

test("the published version is selected by name and cannot downgrade the site", async (t) => {
  const version = publishedVersion(
    JSON.stringify([
      { name: "another-package", version: "9.0.0" },
      { name: library, version: "0.2.0" },
    ])
  );
  assert.equal(version, "0.2.0");
  await assert.rejects(
    prepare(await fixture(t), version, {
      install: async () => assert.fail("must not downgrade"),
    }),
    /older release/
  );
});

const git = async (_command, args) => ({
  stdout: args[0] === "diff" ? files.join("\n") : "a".repeat(40),
});
const commitOptions = {
  execute: git,
  repository: "owner/repo",
  token: "test-token",
};

test("the version commit uses only build inputs and rejects a concurrent main update", async (t) => {
  const root = await fixture(t);
  const before = await inputs(root);
  for (const conflict of [false, true]) {
    let requests = 0;
    const result = commit(root, "0.2.1", {
      ...commitOptions,
      request: async (_url, init) => {
        requests++;
        const input = JSON.parse(init.body).variables.input;
        assert.equal(input.expectedHeadOid, "a".repeat(40));
        assert.equal(input.branch.branchName, "main");
        assert.deepEqual(
          input.fileChanges.additions,
          files.map((path, i) => ({
            path,
            contents: Buffer.from(before[i]).toString("base64"),
          }))
        );
        return {
          ok: true,
          json: async () =>
            conflict
              ? { errors: [{ message: "main advanced" }] }
              : {
                  data: {
                    createCommitOnBranch: {
                      commit: { url: "https://example.com/commit" },
                    },
                  },
                },
        };
      },
    });
    if (conflict) await assert.rejects(result, /main advanced/);
    else assert.equal(await result, true);
    assert.equal(requests, 1);
  }
});

test("unchanged, unrelated or mismatched inputs do not create a commit", async (t) => {
  const root = await fixture(t);
  const options = {
    ...commitOptions,
    request: async () => assert.fail("must not commit"),
  };
  assert.equal(
    await commit(root, "0.2.1", {
      ...options,
      execute: async () => ({ stdout: "" }),
    }),
    false
  );
  await assert.rejects(
    commit(root, "0.2.1", {
      ...options,
      execute: async () => ({ stdout: "src/index.ts" }),
    }),
    /outside the demo pins/
  );
  await assert.rejects(commit(root, "0.3.0", options), /does not match/);
});

test("Changesets bumps the library without moving demo pins to an unpublished version", async (t) => {
  const root = await fixture(t);
  await json(root, "package.json", { name: library, version: "0.2.1" });
  const before = await inputs(root);
  await writeFile(
    join(root, "pnpm-workspace.yaml"),
    await readFile(join(repository, "pnpm-workspace.yaml"))
  );
  const config = JSON.parse(
    await readFile(join(repository, ".changeset/config.json"))
  );
  // Changelog generation can contact GitHub and does not affect dependency versioning.
  config.changelog = false;
  await mkdir(join(root, ".changeset"));
  await json(root, ".changeset/config.json", config);
  await writeFile(
    join(root, ".changeset/release.md"),
    `---\n"${library}": minor\n---\nRelease probe.\n`
  );
  await execute(
    process.execPath,
    [join(repository, "node_modules/@changesets/cli/bin.js"), "version"],
    { cwd: root }
  );
  assert.equal(
    JSON.parse(await readFile(join(root, "package.json"))).version,
    "0.3.0"
  );
  for (const [i, path] of files.slice(0, 2).entries()) {
    assert.deepEqual(
      JSON.parse(await readFile(join(root, path))),
      JSON.parse(before[i])
    );
  }
});
