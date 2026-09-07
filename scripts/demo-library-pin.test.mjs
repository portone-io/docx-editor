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
import { promisify } from "node:util";
import { check, pin } from "./demo-library-pin.mjs";

const library = "@portone/docx-editor";
const json = (path, value) =>
  writeFile(path, `${JSON.stringify(value, null, 2)}\n`);
async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), "demo-library-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await json(join(root, "package.json"), { name: library, version: "0.3.0" });
  for (const dir of ["site", "demo"]) {
    await mkdir(join(root, dir, "node_modules", library), { recursive: true });
    await json(join(root, dir, "package.json"), {
      dependencies: { [library]: "0.2.1" },
    });
  }
  await install(root, "0.2.1");
  return root;
}
async function install(root, version) {
  for (const dir of ["site", "demo"]) {
    await json(join(root, dir, "node_modules", library, "package.json"), {
      name: library,
      version,
    });
  }
  await writeFile(join(root, "pnpm-lock.yaml"), `installed: ${version}\n`);
}
async function inputs(root) {
  return Promise.all(
    ["site/package.json", "demo/package.json", "pnpm-lock.yaml"].map((path) =>
      readFile(join(root, path), "utf8")
    )
  );
}
const noWait = async () => {};

test("the offline check accepts the same inputs before and after a root version bump", async (t) => {
  const root = await fixture(t);
  assert.equal(await check(root), "0.2.1");
  await json(join(root, "package.json"), { name: library, version: "0.4.0" });
  assert.equal(await check(root), "0.2.1");
});

test("the check rejects a workspace link even if the version matches", async (t) => {
  const root = await fixture(t);
  await json(join(root, "package.json"), { name: library, version: "0.2.1" });
  const entry = join(root, "demo/node_modules", library, "package.json");
  await rm(entry);
  await symlink(join(root, "package.json"), entry);
  await assert.rejects(check(root), /working tree/);
});

test("the check rejects disagreeing consumers and a stale installation", async (t) => {
  const root = await fixture(t);
  await json(join(root, "demo/package.json"), {
    dependencies: { [library]: "0.3.0" },
  });
  await assert.rejects(
    check(root),
    /installed library does not match[\s\S]*same version/
  );
});

for (const version of ["workspace:*", "^0.2.1", "0.3.0-beta.1"]) {
  test(`the check rejects ${version}`, async (t) => {
    const root = await fixture(t);
    await json(join(root, "site/package.json"), {
      dependencies: { [library]: version },
    });
    await assert.rejects(check(root), /exact stable version/);
  });
}

test("local preparation resolves latest once and updates both consumers with a mutable CI lockfile", async (t) => {
  const root = await fixture(t);
  const commands = [];
  const result = await pin(root, {
    run: async (command, args) => {
      commands.push([command, args]);
      if (command === "npm") return { stdout: '"0.3.0"' };
      assert.ok(args.includes("--no-frozen-lockfile"));
      assert.ok(args.includes("--ignore-scripts"));
      for (const dir of ["site", "demo"]) {
        const manifest = JSON.parse(
          await readFile(join(root, dir, "package.json"), "utf8")
        );
        assert.equal(manifest.dependencies[library], "0.3.0");
      }
      await install(root, "0.3.0");
    },
  });
  assert.equal(result, "0.3.0");
  assert.equal(await check(root), "0.3.0");
  assert.equal(commands.length, 2);
  assert.equal(commands[0][1][1], `${library}@latest`);
  assert.equal(
    await readFile(join(root, "pnpm-lock.yaml"), "utf8"),
    "installed: 0.3.0\n"
  );
});

test("an explicit publish version waits for visibility and installation without switching to latest", async (t) => {
  const root = await fixture(t);
  let views = 0;
  let installs = 0;
  let waits = 0;
  await pin(root, {
    version: "0.3.0",
    wait: async () => {
      waits++;
    },
    run: async (command, args) => {
      if (command === "npm") {
        assert.equal(args[1], `${library}@0.3.0`);
        if (++views === 1) throw new Error("version not visible yet");
        return { stdout: '"0.3.0"' };
      }
      if (++installs === 1) throw new Error("tarball not visible yet");
      await install(root, "0.3.0");
    },
  });
  assert.equal(views, 2);
  assert.equal(installs, 2);
  assert.equal(waits, 2);
  assert.equal(await check(root), "0.3.0");
});

test("an unavailable release stops after bounded retries without changing inputs", async (t) => {
  const root = await fixture(t);
  const before = await inputs(root);
  let calls = 0;
  await assert.rejects(
    pin(root, {
      version: "0.3.0",
      attempts: 2,
      wait: noWait,
      run: async (command) => {
        assert.equal(command, "npm");
        calls++;
        throw new Error("registry unavailable");
      },
    }),
    /registry unavailable/
  );
  assert.equal(calls, 2);
  assert.deepEqual(await inputs(root), before);
});

test("a failed install restores manifests and lockfile for a later retry", async (t) => {
  const root = await fixture(t);
  const before = await inputs(root);
  await assert.rejects(
    pin(root, {
      version: "0.3.0",
      attempts: 2,
      wait: noWait,
      run: async (command) => {
        if (command === "npm") return { stdout: '"0.3.0"' };
        await writeFile(join(root, "pnpm-lock.yaml"), "partial update");
        throw new Error("download failed");
      },
    }),
    /manifests and lockfile were restored/
  );
  assert.deepEqual(await inputs(root), before);
});

test("an incorrect installed version cannot be reported as a successful update", async (t) => {
  const root = await fixture(t);
  const before = await inputs(root);
  await assert.rejects(
    pin(root, {
      attempts: 1,
      run: async (command) => (command === "npm" ? { stdout: '"0.3.0"' } : {}),
    }),
    /Could not prepare/
  );
  assert.deepEqual(await inputs(root), before);
});

test("the current release needs no reinstall but still checks latest", async (t) => {
  const root = await fixture(t);
  let calls = 0;
  assert.equal(
    await pin(root, {
      run: async (command) => {
        assert.equal(command, "npm");
        calls++;
        return { stdout: '"0.2.1"' };
      },
    }),
    "0.2.1"
  );
  assert.equal(calls, 1);
});

test("the CLI executes from a temporary path, including macOS /var aliases", async (t) => {
  const root = await fixture(t);
  await mkdir(join(root, "scripts"));
  const path = join(root, "scripts/demo-library-pin.mjs");
  await writeFile(
    path,
    await readFile(new URL("./demo-library-pin.mjs", import.meta.url))
  );
  const { stdout } = await promisify(execFile)(process.execPath, [path], {
    cwd: root,
  });
  assert.match(stdout, /@portone\/docx-editor@0\.2\.1 is installed/);
});
