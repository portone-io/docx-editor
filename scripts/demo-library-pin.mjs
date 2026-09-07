import { execFile } from "node:child_process";
import { readFile, realpath, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const repositoryRoot = fileURLToPath(new URL("..", import.meta.url));

const LIBRARY = "@portone/docx-editor";

/**
 * The workspace packages that serve the library to a visitor rather than develop it. They name a
 * released version, so the landing page's demo runs what `npm install` would give the reader and
 * the version badge beside it can be believed.
 */
const DEPENDENTS = ["site", "demo"];

/** A pin, not a range: the badge prints this string, and `v^0.2.1` would be nonsense */
const EXACT_VERSION = /^\d+\.\d+\.\d+(?:-[\w.-]+)?$/;

const manifestPath = (packageDir) =>
  join(repositoryRoot, packageDir, "package.json");

async function readManifest(path) {
  return JSON.parse(await readFile(path, "utf8"));
}

async function readPin(packageDir) {
  const manifest = await readManifest(manifestPath(packageDir));
  return manifest.dependencies?.[LIBRARY];
}

/**
 * Asks the registry for a version, distinguishing "no such version" from "could not ask". A
 * missing answer is the release window, and a missing registry has to be reported rather than
 * read as one.
 */
async function registryVersion(spec) {
  try {
    const { stdout } = await run("npm", ["view", spec, "version"]);
    return { status: "published", version: stdout.trim() };
  } catch (cause) {
    const detail = String(cause.stderr || cause.message || cause);
    if (detail.includes("E404")) return { status: "unpublished" };
    return { status: "unreachable", detail: detail.trim() };
  }
}

/**
 * What the dependent actually loads. An exact pin makes this predictable, so a mismatch means the
 * lockfile was left behind, and a copy that is this repository's own root manifest means the
 * workspace link is back and the demo is running the working tree again.
 */
async function readInstalled(packageDir) {
  const entry = join(
    repositoryRoot,
    packageDir,
    "node_modules",
    ...LIBRARY.split("/"),
    "package.json"
  );
  try {
    const manifest = await readManifest(entry);
    const path = await realpath(entry);
    return {
      version: manifest.version,
      isWorkingTree: path === (await realpath(manifestPath("."))),
    };
  } catch {
    return null;
  }
}

async function check() {
  const problems = [];
  const rootVersion = (await readManifest(manifestPath("."))).version;

  const pins = new Map();
  for (const packageDir of DEPENDENTS) {
    const pin = await readPin(packageDir);
    pins.set(packageDir, pin);
    const where = `${packageDir}/package.json`;

    if (!pin) {
      problems.push(`${where} does not depend on ${LIBRARY}`);
      continue;
    }
    if (!EXACT_VERSION.test(pin)) {
      problems.push(
        `${where} pins ${LIBRARY} to "${pin}", which is a range rather than a released version`
      );
      continue;
    }

    const installed = await readInstalled(packageDir);
    if (!installed) {
      problems.push(
        `${where} pins ${pin}, but nothing is installed at ${packageDir}/node_modules/${LIBRARY}. Run \`pnpm install\`.`
      );
      continue;
    }
    if (installed.version !== pin) {
      problems.push(
        `${where} pins ${pin}, but ${installed.version} is installed. The lockfile is behind the manifest; run \`pnpm install\`.`
      );
    }
    if (installed.isWorkingTree) {
      problems.push(
        `${packageDir} resolves ${LIBRARY} to this repository's own package, so it runs the working tree rather than the released one.`
      );
    }
  }

  const distinct = new Set(pins.values());
  if (distinct.size > 1) {
    const named = DEPENDENTS.map((dir) => `${dir} ${pins.get(dir)}`).join(", ");
    problems.push(
      `the dependents disagree on which version to run (${named}), so the site would build two copies of the library`
    );
  }

  // A pin behind the repository's own version is the release window, and only that: between the
  // release pull request being written and its publish landing, the registry has nothing to move
  // to. Once the registry has that version the pin is simply stale.
  const pin = pins.get(DEPENDENTS[0]);
  const agreed = distinct.size === 1 && pin && EXACT_VERSION.test(pin);
  if (agreed && pin !== rootVersion) {
    const released = await registryVersion(`${LIBRARY}@${rootVersion}`);
    if (released.status === "published") {
      problems.push(
        `${LIBRARY}@${rootVersion} is released, but the demo still runs ${pin}. Run \`pnpm pin:demo-library\` and commit the result.`
      );
    }
    if (released.status === "unreachable") {
      problems.push(
        `the demo runs ${pin} while this repository is at ${rootVersion}, and the registry could not be asked whether ${rootVersion} is released yet:\n${released.detail}`
      );
    }
  }

  if (problems.length > 0) {
    process.stderr.write(
      `The demo does not run the version the site advertises.\n\n${problems
        .map((problem) => `  - ${problem}`)
        .join("\n")}\n`
    );
    process.exitCode = 1;
    return;
  }

  process.stdout.write(
    `${LIBRARY}@${pin} is what the demo runs and what the badge names.\n`
  );
}

async function write() {
  const latest = await registryVersion(LIBRARY);
  if (latest.status !== "published") {
    const detail =
      latest.status === "unreachable" ? `:\n${latest.detail}` : ".";
    process.stderr.write(
      `Could not ask the registry for ${LIBRARY}${detail}\n`
    );
    process.exitCode = 1;
    return;
  }

  for (const packageDir of DEPENDENTS) {
    const path = manifestPath(packageDir);
    const before = await readFile(path, "utf8");
    const after = before.replace(
      new RegExp(`("${LIBRARY}":\\s*)"[^"]*"`),
      `$1"${latest.version}"`
    );
    if (after === before) {
      process.stdout.write(`  ${packageDir} already runs ${latest.version}\n`);
      continue;
    }
    await writeFile(path, after);
    process.stdout.write(`  ${packageDir} now runs ${latest.version}\n`);
  }
}

const wanted = process.argv.includes("--write") ? write : check;
await wanted();
