import { execFile } from "node:child_process";
import { readFile, realpath } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const root = fileURLToPath(new URL("..", import.meta.url));
const library = "@portone/docx-editor";
const branch = "release/next";
const base = "main";

export const tagOf = (version) => `${library}@${version}`;
const titleOf = (version) => `chore: release ${version}`;

/** The CHANGELOG entry for one version, without its heading. */
export function releaseNotes(changelog, version) {
  const lines = changelog.split("\n");
  const start = lines.findIndex((line) => line.trim() === `## ${version}`);
  if (start === -1) throw new Error(`CHANGELOG.md has no entry for ${version}`);
  const rest = lines.slice(start + 1);
  const end = rest.findIndex((line) => /^## /.test(line));
  return `${rest
    .slice(0, end === -1 ? rest.length : end)
    .join("\n")
    .trim()}\n`;
}

function shell(execute, directory, env) {
  return async (command, args, options = {}) => {
    const { stdout } = await execute(command, args, {
      cwd: directory,
      env,
      ...options,
    });
    return stdout.trim();
  };
}

async function signedIn(sh) {
  try {
    await sh("gh", ["auth", "status"]);
  } catch (error) {
    throw new Error("Sign in to GitHub first: gh auth login", { cause: error });
  }
}

/**
 * Turns the changesets pending on main into a version bump and CHANGELOG entry on
 * a pull request opened as the person running this, so its checks start at once.
 * Its title and body both name the version merging it publishes. The branch name
 * is fixed, so running this again rewrites the same pull request, title included,
 * because a rerun can arrive at a different version.
 * Merging the pull request is what publishes.
 */
export async function proposeRelease(
  directory,
  { execute = run, env = process.env } = {}
) {
  const sh = shell(execute, directory, env);
  if (await sh("git", ["status", "--porcelain", "--untracked-files=no"]))
    throw new Error(
      "The working tree has uncommitted changes; commit or stash them first"
    );
  await signedIn(sh);
  await sh("git", ["fetch", "origin", base]);
  const pending = (
    await sh("git", ["ls-tree", "--name-only", `origin/${base}`, ".changeset/"])
  )
    .split("\n")
    .filter((path) => path.endsWith(".md") && !path.endsWith("/README.md"));
  if (pending.length === 0)
    throw new Error(
      `No changeset is pending on origin/${base}; there is nothing to release`
    );
  const previous = await sh("git", ["rev-parse", "--abbrev-ref", "HEAD"]);
  const back =
    previous === "HEAD" ? await sh("git", ["rev-parse", "HEAD"]) : previous;
  await sh("git", ["checkout", "-B", branch, `origin/${base}`]);
  // The changelog writer looks each pull request up on GitHub.
  const token = env.GITHUB_TOKEN || (await sh("gh", ["auth", "token"]));
  await sh("pnpm", ["changeset", "version"], {
    env: { ...env, GITHUB_TOKEN: token },
  });
  const { version } = JSON.parse(
    await readFile(join(directory, "package.json"), "utf8")
  );
  const tag = tagOf(version);
  const title = titleOf(version);
  // Only tracked files move: the bump, the CHANGELOG, and the consumed changesets.
  await sh("git", ["add", "--update"]);
  await sh("git", ["commit", "--message", title, "--message", tag]);
  await sh("git", [
    "push",
    "--force-with-lease",
    "--set-upstream",
    "origin",
    branch,
  ]);
  const body = [
    `Merging this publishes \`${tag}\`.`,
    releaseNotes(
      await readFile(join(directory, "CHANGELOG.md"), "utf8"),
      version
    ).trim(),
  ].join("\n\n");
  const [open] = JSON.parse(
    await sh("gh", [
      "pr",
      "list",
      "--state",
      "open",
      "--head",
      branch,
      "--base",
      base,
      "--json",
      "number,url",
      "--limit",
      "1",
    ])
  );
  let url;
  if (open) {
    await sh("gh", [
      "pr",
      "edit",
      String(open.number),
      "--title",
      title,
      "--body",
      body,
    ]);
    url = open.url;
  } else {
    url = await sh("gh", [
      "pr",
      "create",
      "--base",
      base,
      "--head",
      branch,
      "--title",
      title,
      "--body",
      body,
    ]);
  }
  await sh("git", ["checkout", back]);
  return { version, url };
}

if (
  process.argv[1] &&
  (await realpath(process.argv[1]).catch(() => null)) ===
    fileURLToPath(import.meta.url)
) {
  const [, , command, argument] = process.argv;
  try {
    if (command === "pr") {
      const { version, url } = await proposeRelease(root);
      process.stdout.write(
        `${url}\nMerging this pull request publishes ${tagOf(version)}.\n`
      );
    } else if (command === "notes" && argument) {
      process.stdout.write(
        releaseNotes(
          await readFile(join(root, "CHANGELOG.md"), "utf8"),
          argument
        )
      );
    } else
      throw new Error("Usage: node scripts/release.mjs pr|notes <version>");
  } catch (error) {
    process.stderr.write(`${error.message}\n${error.cause?.message ?? ""}\n`);
    process.exitCode = 1;
  }
}
