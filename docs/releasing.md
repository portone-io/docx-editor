# Releasing

`@portone/docx-editor` is published by [the release workflow](../.github/workflows/release.yml) when a release pull request merges.
A release starts with one command and one merge, and requires npm approval before it goes live.

## How a change becomes a release

A pull request that changes the published package carries a changeset, as [CONTRIBUTING.md](../CONTRIBUTING.md#changesets) explains.
Changesets accumulate on `main` until someone cuts a release:

1. `pnpm release:pr` turns every changeset pending on `main` into a version bump and a CHANGELOG entry on a pull request from `release/next`.
   Its title names the version it stages on merge, and its description opens with that and carries the CHANGELOG entry.
   You open it, so its checks start at once.
   If more changesets land before it merges, run the command again and the same pull request is rewritten.
2. Review and merge it to stage the package.
   Everything on `main` ships with it: a fix cannot go out alone while unreleased work sits ahead of it.

The command needs a signed-in `gh` and refuses a working tree with uncommitted changes or a `main` with no changeset pending.
It refreshes and prunes origin's remote-tracking branches before preparing the release, so a previous release branch's update or deletion does not leave a stale push lease.
The changelog writer looks each pull request up on GitHub, so the command passes the `gh` token along; set `GITHUB_TOKEN` to use another.

## What the workflow checks before publishing

The merge changes `CHANGELOG.md`, which only `pnpm release:pr` writes, so that change is what starts the workflow and nothing else on `main` does.
It stages the version `package.json` declares unless npm already has it.
Before staging it waits for that commit's CI run to pass, because approval on npm is a maintainer's review, not a test run.
Every check a pull request gets guards a release too, the package verification, the real-browser suite, and the workflow audit included.

A maintainer reviews and approves the staged version on npmjs.com or with `npm stage approve <stage-id>` (2FA required).
After approval, rerun all jobs of that exact **Release** run.
The workflow verifies the version on the public registry before tagging the commit, opening the GitHub release from the CHANGELOG entry, and rebuilding the site.
The subsequent site build installs that public version, never a pending one.

## Retrying

A rerun before approval never stages a second copy.
Check npm if a staging attempt failed or its result is ambiguous.
If it is pending, approve it and rerun that exact release run.
If CI failed, rerun CI first.
If npm shows no pending copy and the failed release commit is still `main`, start **Release** manually on `main` with **Retry staging** enabled.
After npm approval, rerun the exact originating run to retain the release commit even if `main` has advanced.

## The version the site's demo runs

On `main` the site and the demo run the editor from this repository's sources, so nothing there names a release and nothing there moves after one.
After publishing succeeds, [Update site release](../.github/workflows/site-release.yml) checks out the release commit, installs the published version into the site and the demo, builds the site, and writes the result to the `production` branch, which Vercel serves.
See [automatic site updates and recovery](../site/README.md#automatic-updates-after-publishing).

## Tags and GitHub releases

Publishing tags the commit and opens a GitHub release from the CHANGELOG entry for that version, once npm has it, so a failed publish leaves no tag behind.
The tag is `@portone/docx-editor@<version>`, the name Changesets builds for a package inside a workspace, and the release carries the same name.

## npm

Staging authenticates through npm [trusted publishing](https://docs.npmjs.com/trusted-publishers), so there is no npm token here to look for.
It needs a trusted publisher configured for the package on npmjs.com, pointing at this repository and `release.yml` with permission to stage publishes, and it only works from a GitHub-hosted runner.
