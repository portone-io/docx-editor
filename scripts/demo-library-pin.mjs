import { execFile } from "node:child_process";
import { readFile, realpath, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const exec = promisify(execFile);
const repositoryRoot = fileURLToPath(new URL("..", import.meta.url));
const LIBRARY = "@portone/docx-editor";
const DEPENDENTS = ["site", "demo"];
const STABLE_VERSION = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/;

const readJson = async (path) => JSON.parse(await readFile(path, "utf8"));

/** Checks the installed build inputs without consulting mutable registry state. */
export async function check(root = repositoryRoot) {
  const problems = [];
  const pins = [];
  const workingTree = await realpath(join(root, "package.json"));
  for (const dir of DEPENDENTS) {
    const manifest = await readJson(join(root, dir, "package.json"));
    const pin = manifest.dependencies?.[LIBRARY];
    pins.push(pin);
    if (typeof pin !== "string" || !STABLE_VERSION.test(pin)) {
      problems.push(`${dir} must pin ${LIBRARY} to an exact stable version`);
      continue;
    }
    const entry = join(root, dir, "node_modules", LIBRARY, "package.json");
    try {
      const installed = await readJson(entry);
      if ((await realpath(entry)) === workingTree) {
        problems.push(`${dir} resolves the library to the working tree`);
      } else if (installed.name !== LIBRARY || installed.version !== pin) {
        problems.push(
          `${dir} pins ${pin}, but its installed library does not match`
        );
      }
    } catch {
      problems.push(
        `${dir} has no readable installed library; run pnpm install`
      );
    }
  }
  if (new Set(pins).size !== 1)
    problems.push("site and demo must use the same version");
  if (problems.length) throw new Error(problems.join("\n"));
  return pins[0];
}

/** Publication and registry reads can become visible at different times. */
async function retry(operation, { wait, attempts }) {
  for (let attempt = 1; ; attempt++) {
    try {
      return await operation();
    } catch (cause) {
      if (attempt >= attempts) throw cause;
      await wait(5000);
    }
  }
}

/** Resolves once, then installs that exact version for both consumers and the badge. */
export async function pin(
  root = repositoryRoot,
  { version = "latest", run = exec, wait = setTimeout, attempts = 6 } = {}
) {
  if (version !== "latest" && !STABLE_VERSION.test(version)) {
    throw new Error(
      "Use latest or an exact stable version (for example, 0.3.0)"
    );
  }
  const policy = { wait, attempts };
  const resolved = await retry(async () => {
    const { stdout } = await run(
      "npm",
      [
        "view",
        `${LIBRARY}@${version}`,
        "version",
        "--json",
        "--registry=https://registry.npmjs.org",
        "--fetch-retries=0",
      ],
      { cwd: root, timeout: 15000 }
    );
    const found = JSON.parse(stdout);
    if (
      typeof found !== "string" ||
      !STABLE_VERSION.test(found) ||
      (version !== "latest" && found !== version)
    ) {
      throw new Error(
        `The registry did not return the requested stable version: ${stdout.trim()}`
      );
    }
    return found;
  }, policy);

  try {
    if ((await check(root)) === resolved) return resolved;
  } catch {
    // A missing installation or a changed pin needs the same install path.
  }

  const paths = [
    ...DEPENDENTS.map((dir) => join(root, dir, "package.json")),
    join(root, "pnpm-lock.yaml"),
  ];
  const originals = await Promise.all(
    paths.map(async (path) => {
      try {
        return await readFile(path, "utf8");
      } catch (cause) {
        if (cause.code === "ENOENT") return null;
        throw cause;
      }
    })
  );
  try {
    for (let i = 0; i < DEPENDENTS.length; i++) {
      const manifest = JSON.parse(originals[i]);
      if (!manifest.dependencies?.[LIBRARY])
        throw new Error(`${paths[i]} has no library dependency`);
      manifest.dependencies[LIBRARY] = resolved;
      await writeFile(paths[i], `${JSON.stringify(manifest, null, 2)}\n`);
    }
    // This command intentionally updates the lockfile, including in CI. Ordinary installs stay frozen.
    await retry(
      () =>
        run("pnpm", ["install", "--no-frozen-lockfile", "--ignore-scripts"], {
          cwd: root,
          timeout: 120000,
        }),
      policy
    );
    await check(root);
    return resolved;
  } catch (cause) {
    await Promise.all(
      paths.map((path, i) =>
        originals[i] === null
          ? rm(path, { force: true })
          : writeFile(path, originals[i])
      )
    );
    throw new Error(
      "Could not prepare the released demo; manifests and lockfile were restored. Run pnpm install before retrying.",
      { cause }
    );
  }
}

if (
  process.argv[1] &&
  (await realpath(process.argv[1]).catch(() => null)) ===
    fileURLToPath(import.meta.url)
) {
  try {
    const args = process.argv.slice(2);
    if (args.length && (args[0] !== "--write" || args.length > 2)) {
      throw new Error(
        "Usage: node scripts/demo-library-pin.mjs [--write [latest|VERSION]]"
      );
    }
    const version =
      args[0] === "--write"
        ? await pin(repositoryRoot, { version: args[1] })
        : await check();
    process.stdout.write(
      `${LIBRARY}@${version} is installed for the demo and its badge.\n`
    );
  } catch (error) {
    process.stderr.write(`${error.message}\n${error.cause?.message ?? ""}\n`);
    process.exitCode = 1;
  }
}
