# Contributing

Thank you for taking the time to improve `@portone-io/docx-editor`. Bug reports, documentation improvements, fixes, and focused feature proposals are welcome.

If you use an AI coding tool, also follow the repository guidance in [AGENTS.md](./AGENTS.md).

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

`pnpm dev:site` serves the `site/` package, the Next.js documentation and landing site, and `pnpm build:site` produces its build. Its landing page mounts the same demo component.

## OOXML changes

For changes that interpret or write OOXML, follow the [specification workflow](./spec/README.md#working-with-the-specification).

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

[Testing](./docs/testing.md) explains the scope and prerequisites of every check.

By contributing, you agree that your contribution is licensed under the Apache License 2.0 used by this project.
