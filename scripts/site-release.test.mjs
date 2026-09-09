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
import { prepare, publish } from "./site-release.mjs";

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

const head = "a".repeat(40);
const git = async (_command, args) => ({
  stdout: args[0] === "diff" ? files.join("\n") : head,
});
const repo = "repos/owner/repo";

/** Answers GitHub as if production served `served` (absent when null) and returns the calls made. */
function github({
  served = null,
  exists = false,
  rejected = null,
  mainPin = null,
  proposed = null,
} = {}) {
  const calls = [];
  const reply = (status, payload) => ({
    ok: status < 400,
    status,
    text: async () => JSON.stringify(payload),
  });
  const request = async (url, init) => {
    const path = url.replace("https://api.github.com/", "");
    const body = init.body ? JSON.parse(init.body) : undefined;
    calls.push({ method: init.method, path, body });
    if (path.startsWith(`${repo}/contents/site/package.json?ref=production`))
      return served
        ? reply(200, { dependencies: { [library]: served } })
        : reply(404, { message: "Not Found" });
    if (path.startsWith(`${repo}/contents/site/package.json?ref=main`))
      return mainPin
        ? reply(200, { dependencies: { [library]: mainPin } })
        : reply(404, { message: "Not Found" });
    if (path === `${repo}/pulls?head=owner:production&base=main&state=open`)
      return reply(200, proposed ? [{ html_url: proposed }] : []);
    if (path === `${repo}/pulls`)
      return reply(201, { html_url: "https://example.com/pull/1" });
    if (path === `${repo}/git/ref/heads/production`)
      return exists ? reply(200, {}) : reply(404, { message: "Not Found" });
    if (
      path === `${repo}/git/refs` ||
      path === `${repo}/git/refs/heads/production`
    )
      return reply(init.method === "POST" ? 201 : 200, {});
    if (path === "graphql")
      return rejected
        ? reply(200, { errors: [{ message: rejected }] })
        : reply(200, {
            data: {
              createCommitOnBranch: {
                commit: { url: "https://example.com/commit" },
              },
            },
          });
    throw new Error(`Unexpected request ${init.method} ${url}`);
  };
  return { calls, request };
}
const publishOptions = (remote, execute = git) => ({
  execute,
  request: remote.request,
  repository: "owner/repo",
  token: "test-token",
});

test("production is rebuilt as the release commit plus the verified inputs", async (t) => {
  const root = await fixture(t);
  const before = await inputs(root);
  for (const exists of [false, true]) {
    const remote = github({ exists });
    assert.equal(await publish(root, "0.2.1", publishOptions(remote)), true);
    const [, , moved, committed] = remote.calls;
    assert.deepEqual(
      remote.calls.map((call) => `${call.method} ${call.path}`),
      [
        `GET ${repo}/contents/site/package.json?ref=production`,
        `GET ${repo}/git/ref/heads/production`,
        exists
          ? `PATCH ${repo}/git/refs/heads/production`
          : `POST ${repo}/git/refs`,
        "POST graphql",
        `GET ${repo}/contents/site/package.json?ref=main`,
        `GET ${repo}/pulls?head=owner:production&base=main&state=open`,
        `POST ${repo}/pulls`,
      ]
    );
    assert.deepEqual(
      moved.body,
      exists
        ? { sha: head, force: true }
        : { ref: "refs/heads/production", sha: head }
    );
    const input = committed.body.variables.input;
    assert.equal(input.branch.branchName, "production");
    assert.equal(input.expectedHeadOid, head);
    assert.deepEqual(
      input.fileChanges.additions,
      files.map((path, i) => ({
        path,
        contents: Buffer.from(before[i]).toString("base64"),
      }))
    );
  }
});

test("a retried release cannot downgrade what production serves", async (t) => {
  const remote = github({ served: "0.3.0", exists: true });
  await assert.rejects(
    publish(await fixture(t), "0.2.1", publishOptions(remote)),
    /older release/
  );
  assert.equal(remote.calls.length, 1);
});

test("unchanged inputs move production to the release commit without a commit", async (t) => {
  const remote = github({ exists: true });
  assert.equal(
    await publish(
      await fixture(t),
      "0.2.1",
      publishOptions(remote, async (_command, args) => ({
        stdout: args[0] === "diff" ? "" : head,
      }))
    ),
    false
  );
  assert.deepEqual(
    remote.calls.map((call) => call.method),
    ["GET", "GET", "PATCH", "GET", "GET", "POST"]
  );
});

const pulls = (remote) =>
  remote.calls.filter((call) => call.path.startsWith(`${repo}/pulls`));

test("the release is proposed on main from the production branch", async (t) => {
  const remote = github({ exists: true });
  await publish(await fixture(t), "0.2.1", publishOptions(remote));
  const [, opened] = pulls(remote);
  assert.equal(`${opened.method} ${opened.path}`, `POST ${repo}/pulls`);
  assert.deepEqual(
    [opened.body.head, opened.body.base],
    ["production", "main"]
  );
  assert.match(opened.body.body, /already runs 0\.2\.1/);
});

test("an open proposal is reported instead of opened again", async (t) => {
  const remote = github({ exists: true, proposed: "https://example.com/7" });
  await publish(await fixture(t), "0.2.1", publishOptions(remote));
  assert.deepEqual(
    pulls(remote).map((call) => call.method),
    ["GET"]
  );
});

test("nothing is proposed when main already pins the release", async (t) => {
  const remote = github({ exists: true, mainPin: "0.2.1" });
  await publish(await fixture(t), "0.2.1", publishOptions(remote));
  assert.deepEqual(pulls(remote), []);
});

test("unrelated or mismatched inputs stop before GitHub is touched", async (t) => {
  const root = await fixture(t);
  const remote = github();
  await assert.rejects(
    publish(
      root,
      "0.2.1",
      publishOptions(remote, async () => ({ stdout: "src/index.ts" }))
    ),
    /outside the demo pins/
  );
  await assert.rejects(
    publish(root, "0.3.0", publishOptions(remote)),
    /does not match/
  );
  assert.equal(remote.calls.length, 0);
});

test("a rejected commit reports GitHub's reason", async (t) => {
  const remote = github({ rejected: "Protected branch update failed" });
  await assert.rejects(
    publish(await fixture(t), "0.2.1", publishOptions(remote)),
    /production \(200\): Protected branch update failed/
  );
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
