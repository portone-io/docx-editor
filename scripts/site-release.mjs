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
const branch = "production";
const api = "https://api.github.com";

function versionParts(version) {
  if (typeof version !== "string" || !stable.test(version)) {
    throw new Error("The site release needs an exact stable version");
  }
  return version.split(".").map(BigInt);
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

/** Installs the release into the checked-out release commit and builds the site from it. */
export async function prepare(
  directory,
  version,
  { install = pin, execute = run } = {}
) {
  versionParts(version);
  await install(directory, { version });
  await execute("pnpm", ["build:site"], { cwd: directory });
}

function github(request, repository, token) {
  const parse = (text) => {
    try {
      return JSON.parse(text);
    } catch {
      return null;
    }
  };
  async function send(url, method, body, accept) {
    const response = await request(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: accept,
        "Content-Type": "application/json",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: AbortSignal.timeout(30000),
    });
    return {
      ok: response.ok,
      status: response.status,
      text: await response.text(),
    };
  }
  return {
    /** A 404 on a read means the branch or file is absent; anything else failing is an error. */
    async rest(method, path, body, accept = "application/vnd.github+json") {
      const result = await send(
        `${api}/repos/${repository}/${path}`,
        method,
        body,
        accept
      );
      if (!result.ok && !(method === "GET" && result.status === 404))
        throw new Error(
          `GitHub ${method} ${path} failed with ${result.status}: ${result.text}`
        );
      return result;
    },
    async createCommit(input) {
      const { ok, status, text } = await send(
        `${api}/graphql`,
        "POST",
        {
          query:
            "mutation($input: CreateCommitOnBranchInput!) { createCommitOnBranch(input: $input) { commit { url } } }",
          variables: { input },
        },
        "application/vnd.github+json"
      );
      const result = ok ? parse(text) : null;
      const errors = result?.errors?.map((error) => error.message) ?? [];
      const url = result?.data?.createCommitOnBranch?.commit?.url;
      if (!ok || errors.length || !url)
        throw new Error(
          `Could not commit the site update to ${branch} (${status}): ${errors.join("; ") || text}`
        );
      return url;
    },
  };
}

/**
 * Rebuilds the production branch as the release commit plus the verified pins.
 * A delayed or manually retried release must not downgrade what production serves.
 */
export async function publish(
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
  const installed = await check(directory);
  if (installed.source !== "npm" || installed.version !== version)
    throw new Error("The installed demo does not match the release to commit");
  const { stdout: changed } = await execute(
    "git",
    ["diff", "--name-only", "HEAD"],
    { cwd: directory }
  );
  const paths = changed.trim().split("\n").filter(Boolean);
  if (paths.some((path) => !files.includes(path)))
    throw new Error(
      "The site update changed files outside the demo pins and lockfile"
    );
  if (!repository || !token)
    throw new Error(
      "GITHUB_REPOSITORY and GH_TOKEN are required to publish the site update"
    );
  const remote = github(request, repository, token);
  const served = await remote.rest(
    "GET",
    `contents/site/package.json?ref=${branch}`,
    undefined,
    "application/vnd.github.raw+json"
  );
  if (served.status !== 404)
    assertNotOlder(version, JSON.parse(served.text).dependencies[library]);
  const { stdout: head } = await execute("git", ["rev-parse", "HEAD"], {
    cwd: directory,
  });
  const sha = head.trim();
  const existing = await remote.rest("GET", `git/ref/heads/${branch}`);
  if (existing.status === 404)
    await remote.rest("POST", "git/refs", { ref: `refs/heads/${branch}`, sha });
  else
    await remote.rest("PATCH", `git/refs/heads/${branch}`, {
      sha,
      force: true,
    });
  const additions = await Promise.all(
    paths.map(async (path) => ({
      path,
      contents: (await readFile(join(directory, path))).toString("base64"),
    }))
  );
  const url = await remote.createCommit({
    branch: { repositoryNameWithOwner: repository, branchName: branch },
    expectedHeadOid: sha,
    message: { headline: `chore: update site demo to ${version}` },
    fileChanges: { additions },
  });
  process.stdout.write(`${url}\n`);
}

if (
  process.argv[1] &&
  (await realpath(process.argv[1]).catch(() => null)) ===
    fileURLToPath(import.meta.url)
) {
  try {
    const version = process.env.SITE_RELEASE_VERSION;
    if (process.argv[2] === "prepare") await prepare(root, version);
    else if (process.argv[2] === "publish") await publish(root, version);
    else
      throw new Error("Usage: node scripts/site-release.mjs prepare|publish");
  } catch (error) {
    process.stderr.write(`${error.message}\n${error.cause?.message ?? ""}\n`);
    process.exitCode = 1;
  }
}
