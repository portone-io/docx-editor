# Releasing

`@portone/docx-editor` is published by [the release workflow](../.github/workflows/release.yml) when a release pull request merges.
A release is one command and one merge.

## How a change becomes a release

A pull request that changes the published package carries a changeset, as [CONTRIBUTING.md](../CONTRIBUTING.md#changesets) explains.
Changesets accumulate on `main` until someone cuts a release:

1. `pnpm release:pr` turns every changeset pending on `main` into a version bump and a CHANGELOG entry on a `chore: release` pull request from `release/next`.
   You open it, so its checks start at once.
   If more changesets land before it merges, run the command again and the same pull request is rewritten.
2. Review and merge it.
   Everything on `main` ships with it: a fix cannot go out alone while unreleased work sits ahead of it.

The command needs a signed-in `gh` and refuses a working tree with uncommitted changes or a `main` with no changeset pending.
The changelog writer looks each pull request up on GitHub, so the command passes the `gh` token along; set `GITHUB_TOKEN` to use another.

## What the workflow checks before publishing

The merge changes `CHANGELOG.md`, which only `pnpm release:pr` writes, so that change is what starts the workflow and nothing else on `main` does.
It publishes the version `package.json` declares unless npm already has it.
Before publishing it waits for that commit's CI run to pass, because nothing downstream can catch a bad tarball once the registry has it.
Every check a pull request gets guards a release too, the package verification, the real-browser suite, and the workflow audit included.

Only then does it publish, tag the commit, open the GitHub release from the CHANGELOG entry, and rebuild the site.

## Retrying

Rerun the failed run from GitHub Actions, or start **Release** by hand on `main`.
It publishes only what npm lacks and opens the release only when it is missing, so a rerun after a failure further down publishes nothing twice.
If CI failed on the commit itself, re-run CI first; the workflow waits on its result.

## The version the site's demo runs

Changesets preserves the published demo pins during release preparation through `bumpVersionsWithWorkspaceProtocolOnly`.
After publishing succeeds, [Update site release](../.github/workflows/site-release.yml) rebuilds the `production` branch, which Vercel serves, from the release commit and the published version.
It then opens a pull request from `production` to `main`, unless `main` already pins that version, so the `main` preview and local site builds follow the release once it is merged.
A workflow token opens that pull request, so its checks wait for **Approve workflows to run** in the merge box, the one approval a release still asks of a person.
See [automatic site updates and recovery](../site/README.md#automatic-updates-after-publishing).

## Tags and GitHub releases

Publishing tags the commit and opens a GitHub release from the CHANGELOG entry for that version, once npm has it, so a failed publish leaves no tag behind.
The tag is `@portone/docx-editor@<version>`, the name Changesets builds for a package inside a workspace, and the release carries the same name.

## npm

Publishing authenticates through npm [trusted publishing](https://docs.npmjs.com/trusted-publishers), so there is no npm token here to look for.
It needs a trusted publisher configured for the package on npmjs.com, pointing at this repository and `release.yml`, and it only works from a GitHub-hosted runner.
