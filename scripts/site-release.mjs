import { execFile } from "node:child_process";
import { readFile, realpath } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { check, pin } from "./demo-library-pin.mjs";

const run = promisify(execFile);
const root = fileURLToPath(new URL("..", import.meta.url));
const library = "@portone/docx-editor";
const files = ["site/package.json", "demo/package.json", "pnpm-lock.yaml"];
const stable = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/;

function versionParts(version) {
  if (typeof version !== "string" || !stable.test(version)) {
    throw new Error("The site release needs an exact stable version");
  }
  return version.split(".").map(BigInt);
}

export function publishedVersion(output) {
  const matches = JSON.parse(output).filter((pkg) => pkg.name === library);
  if (matches.length !== 1)
    throw new Error("Expected one published DOCX editor package");
  versionParts(matches[0].version);
  return matches[0].version;
}

export function assertNotOlder(version, current) {
  const next = versionParts(version);
  const previous = versionParts(current);
  for (let i = 0; i < next.length; i++) {
    if (next[i] > previous[i]) return;
    if (next[i] < previous[i])
      throw new Error(
        `Refusing to replace demo ${current} with older release ${version}`
      );
  }
}

/** A delayed or manually retried release must not downgrade a newer site's pin. */
export async function prepare(
  directory,
  version,
  { install = pin, execute = run } = {}
) {
  const manifest = JSON.parse(
    await readFile(join(directory, "site/package.json"), "utf8")
  );
  assertNotOlder(version, manifest.dependencies[library]);
  await install(directory, { version });
  await execute("pnpm", ["build:site"], { cwd: directory });
}

/** Atomically updates only the validated inputs, and refuses if main advanced during the build. */
export async function commit(
  directory,
  version,
  {
    execute = run,
    request = fetch,
    repository = process.env.GITHUB_REPOSITORY,
    token = process.env.GH_TOKEN,
  } = {}
) {
  versionParts(version);
  if ((await check(directory)) !== version)
    throw new Error("The installed demo does not match the release to commit");
  const { stdout: changed } = await execute(
    "git",
    ["diff", "--name-only", "HEAD"],
    { cwd: directory }
  );
  const paths = changed.trim().split("\n").filter(Boolean);
  if (!paths.length) return false;
  if (paths.some((path) => !files.includes(path)))
    throw new Error(
      "The site update changed files outside the demo pins and lockfile"
    );
  if (!repository || !token)
    throw new Error(
      "GITHUB_REPOSITORY and GH_TOKEN are required to commit the site update"
    );
  const { stdout: head } = await execute("git", ["rev-parse", "HEAD"], {
    cwd: directory,
  });
  const additions = await Promise.all(
    paths.map(async (path) => ({
      path,
      contents: (await readFile(join(directory, path))).toString("base64"),
    }))
  );
  const response = await request("https://api.github.com/graphql", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      query:
        "mutation($input: CreateCommitOnBranchInput!) { createCommitOnBranch(input: $input) { commit { url } } }",
      variables: {
        input: {
          branch: { repositoryNameWithOwner: repository, branchName: "main" },
          expectedHeadOid: head.trim(),
          message: { headline: `chore: update site demo to ${version}` },
          fileChanges: { additions },
        },
      },
    }),
    signal: AbortSignal.timeout(30000),
  });
  const result = await response.json();
  if (
    !response.ok ||
    result.errors?.length ||
    !result.data?.createCommitOnBranch?.commit?.url
  ) {
    throw new Error(
      "Could not commit the site update. If main advanced, rerun Update site release against current main."
    );
  }
  process.stdout.write(`${result.data.createCommitOnBranch.commit.url}\n`);
  return true;
}

if (
  process.argv[1] &&
  (await realpath(process.argv[1]).catch(() => null)) ===
    fileURLToPath(import.meta.url)
) {
  try {
    const version = process.env.SITE_RELEASE_VERSION;
    if (process.argv[2] === "prepare") await prepare(root, version);
    else if (process.argv[2] === "commit") {
      if (!(await commit(root, version)))
        process.stdout.write(
          "The site already pins this release; retry a failed Vercel deployment in Vercel.\n"
        );
    } else
      throw new Error("Usage: node scripts/site-release.mjs prepare|commit");
  } catch (error) {
    process.stderr.write(`${error.message}\n${error.cause?.message ?? ""}\n`);
    process.exitCode = 1;
  }
}
