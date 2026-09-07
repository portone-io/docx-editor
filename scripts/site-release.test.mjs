import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import {
  assertNotOlder,
  commit,
  prepare,
  publishedVersion,
} from "./site-release.mjs";

const library = "@portone/docx-editor";
const json = (path, value) => writeFile(path, `${JSON.stringify(value)}\n`);
async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), "site-release-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await json(join(root, "package.json"), { name: library, version: "0.4.0" });
  for (const dir of ["site", "demo"]) {
    await mkdir(join(root, dir, "node_modules", library), { recursive: true });
    await json(join(root, dir, "package.json"), {
      dependencies: { [library]: "0.3.0" },
    });
    await json(join(root, dir, "node_modules", library, "package.json"), {
      name: library,
      version: "0.3.0",
    });
  }
  await writeFile(join(root, "pnpm-lock.yaml"), "installed: 0.3.0\n");
  return root;
}
const execute = async (_command, args) => ({
  stdout:
    args[0] === "diff"
      ? "demo/package.json\npnpm-lock.yaml\nsite/package.json\n"
      : `${"a".repeat(40)}\n`,
});
const options = { execute, repository: "owner/repo", token: "test-token" };

test("the published package is selected by name, not array position", () => {
  assert.equal(
    publishedVersion(
      JSON.stringify([
        { name: "another-package", version: "9.0.0" },
        { name: library, version: "0.3.0" },
      ])
    ),
    "0.3.0"
  );
  assert.throws(() => publishedVersion("[]"), /Expected one/);
  assert.throws(
    () =>
      publishedVersion(
        JSON.stringify([{ name: library, version: "0.3.0-beta.1" }])
      ),
    /stable version/
  );
});

test("late release runs cannot downgrade a newer site", () => {
  assert.throws(() => assertNotOlder("0.2.1", "0.3.0"), /older release/);
  assert.throws(() => assertNotOlder("0.9.0", "0.10.0"), /older release/);
  assertNotOlder("0.10.0", "0.9.0");
  assertNotOlder("1.0.0", "0.99.0");
  assertNotOlder("0.3.0", "0.3.0");
});

test("a stale release is refused before installation or building", async (t) => {
  const root = await fixture(t);
  await assert.rejects(
    prepare(root, "0.2.1", {
      install: async () => assert.fail("must not install"),
      execute: async () => assert.fail("must not build"),
    }),
    /older release/
  );
});

test("site building waits for installation to complete", async (t) => {
  const root = await fixture(t);
  const steps = [];
  await prepare(root, "0.3.0", {
    install: async (directory, { version }) => {
      assert.equal(directory, root);
      assert.equal(version, "0.3.0");
      await new Promise((resolve) => setImmediate(resolve));
      steps.push("installed");
    },
    execute: async (command, args) => {
      assert.deepEqual(steps, ["installed"]);
      assert.equal(command, "pnpm");
      assert.deepEqual(args, ["build:site"]);
      steps.push("built");
    },
  });
  assert.deepEqual(steps, ["installed", "built"]);
});

test("installation failure stops before the site build", async (t) => {
  const root = await fixture(t);
  await assert.rejects(
    prepare(root, "0.3.0", {
      install: async () => {
        throw new Error("not yet installable");
      },
      execute: async () => assert.fail("must not build"),
    }),
    /not yet installable/
  );
});

test("a build failure remains a failed site update", async (t) => {
  const root = await fixture(t);
  await assert.rejects(
    prepare(root, "0.3.0", {
      install: async () => {},
      execute: async () => {
        throw new Error("site build failed");
      },
    }),
    /site build failed/
  );
});

test("the version commit contains only verified files and the head that was built", async (t) => {
  const root = await fixture(t);
  let requests = 0;
  assert.equal(
    await commit(root, "0.3.0", {
      ...options,
      request: async (url, init) => {
        requests++;
        assert.equal(url, "https://api.github.com/graphql");
        const input = JSON.parse(init.body).variables.input;
        assert.equal(input.expectedHeadOid, "a".repeat(40));
        assert.deepEqual(input.branch, {
          repositoryNameWithOwner: "owner/repo",
          branchName: "main",
        });
        assert.deepEqual(
          input.fileChanges.additions.map((file) => file.path).sort(),
          ["demo/package.json", "pnpm-lock.yaml", "site/package.json"]
        );
        for (const file of input.fileChanges.additions.filter((file) =>
          file.path.endsWith("package.json")
        )) {
          assert.equal(
            JSON.parse(Buffer.from(file.contents, "base64").toString())
              .dependencies[library],
            "0.3.0"
          );
        }
        return {
          ok: true,
          json: async () => ({
            data: {
              createCommitOnBranch: {
                commit: { url: "https://example.com/commit" },
              },
            },
          }),
        };
      },
    }),
    true
  );
  assert.equal(requests, 1);
});

test("concurrent main updates are reported without retrying a stale commit", async (t) => {
  const root = await fixture(t);
  let requests = 0;
  await assert.rejects(
    commit(root, "0.3.0", {
      ...options,
      request: async () => {
        requests++;
        return {
          ok: true,
          json: async () => ({
            errors: [{ message: "expectedHeadOid does not match" }],
          }),
        };
      },
    }),
    /main advanced/
  );
  assert.equal(requests, 1);
});

test("an unchanged retry does not create a redundant commit", async (t) => {
  const root = await fixture(t);
  assert.equal(
    await commit(root, "0.3.0", {
      ...options,
      execute: async () => ({ stdout: "" }),
      request: async () => assert.fail("must not commit"),
    }),
    false
  );
});

test("unrelated file changes cannot be swept into the version commit", async (t) => {
  const root = await fixture(t);
  await assert.rejects(
    commit(root, "0.3.0", {
      ...options,
      execute: async () => ({ stdout: "src/index.ts\n" }),
      request: async () => assert.fail("must not commit"),
    }),
    /outside the demo pins/
  );
});

test("a mismatched installation cannot be committed as the requested release", async (t) => {
  const root = await fixture(t);
  await assert.rejects(
    commit(root, "0.4.0", {
      ...options,
      request: async () => assert.fail("must not commit"),
    }),
    /does not match/
  );
});
