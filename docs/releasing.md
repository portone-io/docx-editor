# Releasing

`@portone/docx-editor` is released from `main` by [the release workflow](../.github/workflows/release.yml).
Publishing is a merge, not a command.

## How a change becomes a release

A pull request that changes the published package carries a changeset, as [CONTRIBUTING.md](../CONTRIBUTING.md#changesets) explains.

Once one lands on `main`, the workflow opens a `chore: release` pull request holding the version bump and the CHANGELOG entries the pending changesets add up to, and rewrites it as more land.
Merging that pull request publishes, and everything on `main` ships with it: a fix cannot go out alone while unreleased work sits ahead of it.

Its Actions runs wait for approval, since the workflow opened the pull request rather than a person.
**Approve workflows to run**, in the merge box, starts them; every rewrite re-arms it.

## The version the site's demo runs

Changesets preserves the published demo pins during release preparation through `bumpVersionsWithWorkspaceProtocolOnly`. After publishing succeeds, a separate workflow rebuilds the `production` branch, which Vercel serves, from the release commit and the published version.
It then opens a pull request from `production` to `main`, unless `main` already pins that version, so the `main` preview and local site builds follow the release once it is merged.
A workflow token opens it too, so its runs wait for the same **Approve workflows to run** as the release pull request's.
See [automatic site updates and recovery](../site/README.md#automatic-updates-after-publishing).

## What the workflow decides

Every push to `main` runs the workflow, which first works out which of three things this push is.
A pending changeset means a release is being proposed, so it writes the release pull request.
No changeset, and a version the registry has never seen, means that pull request has been merged, so it publishes.
Anything else is an ordinary commit and the run stops there.

The publish path passes through a gate that waits for the commit's CI run to pass before anything reaches npm, because nothing downstream can catch a bad tarball once the registry has it.
Every check a pull request gets guards a release too, the package verification, the real-browser suite, and the workflow audit included.
A gate failure blocks the release, and the version bump stays on `main` until the next push retries it.
To retry sooner, re-run CI on that commit, then run **Release** by hand in GitHub Actions.

## Tags and GitHub releases

Publishing tags the commit and opens a GitHub release from the CHANGELOG entry for that version.
The tag is `@portone/docx-editor@<version>`, the name Changesets builds for a package inside a workspace.

## npm

Publishing authenticates through npm [trusted publishing](https://docs.npmjs.com/trusted-publishers), so there is no npm token here to look for.
It needs a trusted publisher configured for the package on npmjs.com, pointing at this repository and `release.yml`, and it only works from a GitHub-hosted runner.
