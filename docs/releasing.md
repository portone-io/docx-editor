# Releasing

`@portone/docx-editor` is released from `main` by [the release workflow](../.github/workflows/release.yml).
Publishing is a merge, not a command.

## How a change becomes a release

A pull request that changes the published package carries a changeset, as [CONTRIBUTING.md](../CONTRIBUTING.md#changesets) explains.

Once one lands on `main`, the workflow opens a `chore: release` pull request holding the version bump and the CHANGELOG entries the pending changesets add up to, and rewrites it as more land.
Merging that pull request publishes, and everything on `main` ships with it: a fix cannot go out alone while unreleased work sits ahead of it.

Its Actions runs wait for approval, since the workflow opened the pull request rather than a person.
**Approve workflows to run**, in the merge box, starts them; every rewrite re-arms it.

## npm

Publishing authenticates through npm [trusted publishing](https://docs.npmjs.com/trusted-publishers), so there is no npm token here to look for.
It needs a trusted publisher configured for the package on npmjs.com, pointing at this repository and `release.yml`, and it only works from a GitHub-hosted runner.
