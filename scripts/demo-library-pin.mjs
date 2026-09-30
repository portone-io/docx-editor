import { execFile } from "node:child_process";
import { readFile, realpath, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const exec = promisify(execFile);
const repositoryRoot = fileURLToPath(new URL("..", import.meta.url));
export const LIBRARY = "@portone/docx-editor";
const DEPENDENTS = ["site", "demo"];
export const BUILD_INPUTS = [
  ...DEPENDENTS.map((dir) => `${dir}/package.json`),
  "pnpm-lock.yaml",
];
const WORKSPACE = "workspace:*";
export const STABLE_VERSION = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/;

const readJson = async (path) => JSON.parse(await readFile(path, "utf8"));

/**
 * Checks the installed build inputs without consulting mutable registry state.
 * Both consumers run the sources through the workspace link, or both pin the same
 * published release and have it installed; anything in between is an error.
 */
export async function check(root = repositoryRoot) {
  const problems = [];
  const declared = [];
  const workingTree = await realpath(join(root, "package.json"));
  for (const dir of DEPENDENTS) {
    const manifest = await readJson(join(root, dir, "package.json"));
    const range = manifest.dependencies?.[LIBRARY];
    const linked = range === WORKSPACE;
    if (!linked && (typeof range !== "string" || !STABLE_VERSION.test(range))) {
      problems.push(
        `${dir} must depend on ${LIBRARY} through ${WORKSPACE} or an exact stable version`
      );
      continue;
    }
    declared.push(range);
    const entry = join(root, dir, "node_modules", LIBRARY, "package.json");
    try {
      const installed = await readJson(entry);
      const resolvesToSources = (await realpath(entry)) === workingTree;
      if (linked && !resolvesToSources) {
        problems.push(
          `${dir} declares ${WORKSPACE}, but its installed library is not the working tree`
        );
      } else if (!linked && resolvesToSources) {
        problems.push(
          `${dir} pins ${range}, but its installed library resolves to the working tree`
        );
      } else if (
        !linked &&
        (installed.name !== LIBRARY || installed.version !== range)
      ) {
        problems.push(
          `${dir} pins ${range}, but its installed library does not match`
        );
      }
    } catch {
      problems.push(
        `${dir} has no readable installed library; run pnpm install`
      );
    }
  }
  if (new Set(declared).size > 1)
    problems.push("site and demo must depend on the same library");
  if (problems.length) throw new Error(problems.join("\n"));
  return declared[0] === WORKSPACE
    ? { source: "workspace" }
    : { source: "npm", version: declared[0] };
}

// npm and pnpm report a version that is published but not yet propagated with these codes.
const NOT_LISTED_YET =
  /\bE404\b|\bERR_PNPM_(?:NO_MATCHING_VERSION|FETCH_404)\b/;

const duration = (ms) =>
  ms % 60000 === 0 ? `${ms / 60000} min` : `${Math.round(ms / 1000)} s`;

/** Publication and registry reads can become visible minutes apart. */
async function untilListed(
  operation,
  { spec, state, windowMs, recovery, wait, now, log }
) {
  const deadline = now() + windowMs;
  for (let delay = 5000; ; delay = Math.min(delay * 2, 60000)) {
    try {
      return await operation();
    } catch (cause) {
      if (!NOT_LISTED_YET.test(`${cause?.stdout ?? ""}${cause?.stderr ?? ""}`))
        throw cause;
      if (now() + delay > deadline)
        throw new Error(
          `${spec} was not ${state} within ${duration(windowMs)}. Check it with \`npm view ${spec} version\`, then ${recovery}.`,
          { cause }
        );
      log(`${spec} is not ${state} yet; retrying in ${duration(delay)}.`);
      await wait(delay);
    }
  }
}

/** Resolves once, then installs that exact version for both consumers and the badge. */
export async function pin(
  root = repositoryRoot,
  {
    version = "latest",
    registryWaitMs = 30000,
    recovery = "rerun pnpm pin:demo-library",
    run = exec,
    wait = setTimeout,
    now = Date.now,
    log = (line) => process.stderr.write(`${line}\n`),
  } = {}
) {
  if (version !== "latest" && !STABLE_VERSION.test(version)) {
    throw new Error(
      "Use latest or an exact stable version (for example, 0.3.0)"
    );
  }
  const policy = { windowMs: registryWaitMs, recovery, wait, now, log };
  const resolved = await untilListed(
    async () => {
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
    },
    {
      ...policy,
      spec: `${LIBRARY}@${version}`,
      state: "listed on the registry",
    }
  );

  try {
    const current = await check(root);
    if (current.source === "npm" && current.version === resolved)
      return resolved;
  } catch {
    // A missing installation or a changed pin needs the same install path.
  }

  const paths = BUILD_INPUTS.map((path) => join(root, path));
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
    await untilListed(
      () =>
        run("pnpm", ["install", "--no-frozen-lockfile", "--ignore-scripts"], {
          cwd: root,
          timeout: 120000,
        }),
      { ...policy, spec: `${LIBRARY}@${resolved}`, state: "installable" }
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
    const state =
      args[0] === "--write"
        ? {
            source: "npm",
            version: await pin(repositoryRoot, { version: args[1] }),
          }
        : await check();
    process.stdout.write(
      state.source === "workspace"
        ? `${LIBRARY} resolves to the sources for the demo and its badge.\n`
        : `${LIBRARY}@${state.version} is installed for the demo and its badge.\n`
    );
  } catch (error) {
    process.stderr.write(`${error.message}\n${error.cause?.message ?? ""}\n`);
    process.exitCode = 1;
  }
}
