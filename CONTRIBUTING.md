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
```

[Testing](https://github.com/portone-io/docx-editor/blob/main/docs/testing.md) explains the scope and prerequisites of every check.

## Changesets

If you changed anything under `src/`, run `pnpm changeset` and commit the file it writes.
It becomes the CHANGELOG entry, so write it for someone reading release notes rather than the diff.
Choose `patch` unless the change adds to or breaks the public API.
Those are `minor` while the package is below 1.0.

By contributing, you agree that your contribution is licensed under the Apache License 2.0 used by this project.

## Commit messages

Write the package name as a code span (`@portone/docx-editor`) in commit messages, pull request bodies, and comments, and leave the scope out of pull request titles.
Written bare, the scope reads as a mention of the unrelated GitHub user who holds that handle.
