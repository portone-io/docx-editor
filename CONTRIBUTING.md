# Contributing

Thank you for taking the time to improve `@portone/docx-editor`. Bug reports, documentation improvements, fixes, and focused feature proposals are welcome.

<!-- This file is published to npm while the documents it points at are not, so those links are
     absolute rather than repository-relative. -->

If you use an AI coding tool, also follow the repository guidance in [AGENTS.md](https://github.com/portone-io/docx-editor/blob/main/AGENTS.md).

## Reporting bugs

[Open an issue](https://github.com/portone-io/docx-editor/issues/new) with enough detail to reproduce the problem. Include what you expected and what happened instead.

If the problem depends on a particular DOCX file, attach a minimal example after removing confidential or identifying content.

## Setup

```sh
git clone https://github.com/portone-io/docx-editor.git
cd docx-editor
pnpm install
pnpm dev
```

`pnpm dev` serves the `demo/` package with `demo.docx` open and reloads it as the source changes. `pnpm build:demo` produces its static bundle.

`pnpm dev:site` serves the [documentation and landing site](https://github.com/portone-io/docx-editor/blob/main/site/README.md), and `pnpm build:site` produces its build.

## OOXML changes

For changes that interpret or write OOXML, follow the [specification workflow](https://github.com/portone-io/docx-editor/blob/main/spec/README.md#working-with-the-specification).

## Tests

Run the default local checks before opening a pull request:

```sh
pnpm check
```

Run a specialized check when your change affects its area:

```sh
pnpm test:package
pnpm verify:package
pnpm test:e2e
pnpm bench
```

A pull request that changes performance-sensitive code attaches its `pnpm bench` output to the description.

[Testing](https://github.com/portone-io/docx-editor/blob/main/docs/testing.md) explains the scope and prerequisites of every check.
CI runs `pnpm test:package` on every pull request in a separate job; it is not part of the local `pnpm check` command.

A command added to `./commands` or `./table` needs a writer probe in `src/docx/__testing__/writerProbes.ts` with a `check` of its immediate effect, so its export is validated before another command can overwrite it, or an entry in `NOT_A_WRITER` giving the reason it reaches no writer.

## Changesets

If you changed production code under `src/`, run `pnpm changeset` and commit the file it writes.
It becomes the CHANGELOG entry, so write it for someone reading release notes rather than the diff.
Changes limited to tests, fixtures, scripts, or documentation do not need a changeset.

The public API is the exported names and types of the four entry points - `.`, `./core`, `./commands`, and `./table` - together with the parts of the document model that [What a plugin may rely on](https://docx-editor.portone.io/docs/editor-api/plugins-and-presets#what-a-plugin-may-rely-on) calls stable.
An attr that page calls internal is not part of it, and neither is anything under `src/` that no entry point reaches.

Choose `patch` unless the change adds to that surface, takes something out of it, or changes what a part of it means.
Those are `minor` while the package is below 1.0.
A diff to `src/schema/attrRoles.ts` is where to check this: a change to a stable document-model attr is `minor` even when the declaration reports do not change.
A declaration added, removed, or rewritten in `etc/*.api.md` after running `pnpm api:update` is what that looks like; a line that only gains or loses an `(undocumented)` marker is not.
Commit the regenerated report in the same commit as the changeset.

The changesets that have landed become a release the way [Releasing](https://github.com/portone-io/docx-editor/blob/main/docs/releasing.md) describes.

By contributing, you agree that your contribution is licensed under the Apache License 2.0 used by this project.

## Commit messages

Write the package name as a code span (`@portone/docx-editor`) in commit messages, pull request bodies, and comments, and leave the scope out of pull request titles.
Written bare, the scope reads as a mention of the unrelated GitHub user who holds that handle.
