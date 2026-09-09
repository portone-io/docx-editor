import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { proposeRelease, releaseNotes } from "./release.mjs";

const library = "@portone/docx-editor";
const sha = "b".repeat(40);
const changelog = [
  `# ${library}`,
  "",
  "## 0.6.0",
  "",
  "### Minor Changes",
  "",
  "- A new thing.",
  "",
  "## 0.5.0",
  "",
  "### Patch Changes",
  "",
  "- An old thing.",
  "",
].join("\n");

async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), "release-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await writeFile(
    join(root, "package.json"),
    JSON.stringify({ name: library, version: "0.5.0" })
  );
  await writeFile(join(root, "CHANGELOG.md"), `# ${library}\n\n## 0.5.0\n`);
  return root;
}

/**
 * Answers each command from the first matching prefix and records every call.
 * An Error answer is thrown, as a failed command would be.
 */
function terminal(answers) {
  const calls = [];
  const execute = async (command, args, options) => {
    const line = [command, ...args].join(" ");
    calls.push({ line, options });
    const found = Object.entries(answers).find(([prefix]) =>
      line.startsWith(prefix)
    );
    const answer =
      typeof found?.[1] === "function" ? await found[1](options) : found?.[1];
    if (answer instanceof Error) throw answer;
    return { stdout: answer ?? "", stderr: "" };
  };
  const lines = () => calls.map((call) => call.line);
  const option = (prefix, key) =>
    calls.find((call) => call.line.startsWith(prefix))?.options[key];
  return { execute, lines, option };
}

async function bump(options) {
  await writeFile(
    join(options.cwd, "package.json"),
    JSON.stringify({ name: library, version: "0.6.0" })
  );
  await writeFile(join(options.cwd, "CHANGELOG.md"), changelog);
  return "";
}
const proposal = {
  "git ls-tree":
    ".changeset/README.md\n.changeset/config.json\n.changeset/a.md",
  "git rev-parse --abbrev-ref HEAD": "feature/x",
  "gh auth token": "gh-token",
  "pnpm changeset version": bump,
  "gh pr list": "[]",
  "gh pr create": "https://example.com/pull/9",
};
const env = { PATH: "/usr/bin" };

test("release notes are one version's CHANGELOG entry without its heading", () => {
  assert.equal(
    releaseNotes(changelog, "0.6.0"),
    "### Minor Changes\n\n- A new thing.\n"
  );
  assert.equal(
    releaseNotes(changelog, "0.5.0"),
    "### Patch Changes\n\n- An old thing.\n"
  );
  assert.throws(() => releaseNotes(changelog, "0.4.0"), /no entry for 0\.4\.0/);
});

test("the release pull request is cut from origin/main and opened by the person running it", async (t) => {
  const root = await fixture(t);
  const shell = terminal(proposal);
  const result = await proposeRelease(root, { execute: shell.execute, env });
  assert.deepEqual(result, {
    version: "0.6.0",
    url: "https://example.com/pull/9",
  });
  const body = [
    "### Minor Changes",
    "",
    "- A new thing.",
    "",
    `Merging this publishes \`${library}@0.6.0\`.`,
  ].join("\n");
  assert.deepEqual(shell.lines(), [
    "git status --porcelain --untracked-files=no",
    "gh auth status",
    "git fetch origin main",
    "git ls-tree --name-only origin/main .changeset/",
    "git rev-parse --abbrev-ref HEAD",
    "git checkout -B release/next origin/main",
    "gh auth token",
    "pnpm changeset version",
    "git add --update",
    `git commit --message chore: release --message ${library}@0.6.0`,
    "git push --force-with-lease --set-upstream origin release/next",
    "gh pr list --state open --head release/next --base main --json number,url --limit 1",
    `gh pr create --base main --head release/next --title chore: release --body ${body}`,
    "git checkout feature/x",
  ]);
  assert.equal(
    shell.option("pnpm changeset version", "env").GITHUB_TOKEN,
    "gh-token"
  );
  assert.equal(shell.option("git fetch", "env").GITHUB_TOKEN, undefined);
  assert.equal(
    JSON.parse(await readFile(join(root, "package.json"), "utf8")).version,
    "0.6.0"
  );
});

test("a token already in the environment is used and a detached HEAD is returned to by commit", async (t) => {
  const shell = terminal({
    ...proposal,
    "git rev-parse --abbrev-ref HEAD": "HEAD",
    "git rev-parse HEAD": sha,
  });
  await proposeRelease(await fixture(t), {
    execute: shell.execute,
    env: { ...env, GITHUB_TOKEN: "given" },
  });
  assert.ok(!shell.lines().includes("gh auth token"));
  assert.equal(
    shell.option("pnpm changeset version", "env").GITHUB_TOKEN,
    "given"
  );
  assert.equal(shell.lines().at(-1), `git checkout ${sha}`);
});

test("an open release pull request is rewritten rather than opened again", async (t) => {
  const shell = terminal({
    ...proposal,
    "gh pr list": JSON.stringify([
      { number: 9, url: "https://example.com/pull/9" },
    ]),
  });
  const result = await proposeRelease(await fixture(t), {
    execute: shell.execute,
    env,
  });
  assert.equal(result.url, "https://example.com/pull/9");
  const pulls = shell.lines().filter((line) => line.startsWith("gh pr "));
  assert.equal(pulls.length, 2);
  assert.match(pulls[1], /^gh pr edit 9 --body ### Minor Changes/);
});

for (const [reason, answers, message] of [
  ["a dirty working tree", { "git status": " M src/index.ts" }, /uncommitted/],
  [
    "a signed-out gh",
    { "gh auth status": new Error("not logged in") },
    /gh auth login/,
  ],
  [
    "no pending changeset",
    { "git ls-tree": ".changeset/README.md\n.changeset/config.json" },
    /nothing to release/,
  ],
]) {
  test(`${reason} stops the release pull request before any branch moves`, async (t) => {
    const shell = terminal({ ...proposal, ...answers });
    await assert.rejects(
      proposeRelease(await fixture(t), { execute: shell.execute, env }),
      message
    );
    assert.ok(!shell.lines().some((line) => line.startsWith("git checkout")));
  });
}
