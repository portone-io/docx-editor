# AGENTS.md

Working rules for `@portone/docx-editor`, a DOCX editor built on ProseMirror.

## Context routing

Use [README.md](./README.md) for product context and [CONTRIBUTING.md](./CONTRIBUTING.md) for the contribution workflow. Read only the documents relevant to the task.
When adding a reusable guide, add it here with when to read it.
When adding or changing a feature, update the relevant document and revise the surrounding content so its claims remain accurate and ordered by importance.
Keep one source of truth for each concern and link to it instead of repeating its details in another document.

- For an OOXML decision, check [the existing specification notes](./spec/notes/README.md), then follow [the specification guide](./spec/README.md) when no note covers it. Do not load specification prose or schemas for unrelated work.
- Read [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) when adding or replacing redistributed material or its attribution.
- Read [docs/architecture.md](./docs/architecture.md) before moving code, changing folder dependencies, or adding an entry point.
- Read [docs/testing.md](./docs/testing.md) before changing test structure or packaging checks.
- Read [docs/releasing.md](./docs/releasing.md) before changing how a version is decided, published, or announced.
- The user-facing documentation lives in the site's MDX pages, which the site serves; edit them there rather than mirroring them elsewhere.
  - Read [site/content/docs/features.mdx](./site/content/docs/features.mdx) for a support claim or a behavior change.
  - Read [site/content/docs/core.mdx](./site/content/docs/core.mdx) for programmatic DOCX import and export, including verifying a commenter's file on the server.
  - Read [site/content/docs/custom-controls.mdx](./site/content/docs/custom-controls.mdx) for commands, queries, plugins, custom controls, or the comment ownership rule.
  - Read [site/content/docs/props.mdx](./site/content/docs/props.mdx) for the `mode` union.
- Read [site/README.md](./site/README.md) for how the documentation and landing site is built and run.
- Read [e2e/README.md](./e2e/README.md) only for a real-browser, IME, focus, or keyboard task.
- Read [__fixtures__/README.md](./__fixtures__/README.md) only when adding, replacing, or interpreting a fixture document.

## Conventions

### Language

Use English for source code, comments, test descriptions, test data, and public documentation. Other languages are appropriate when the language itself is required to exercise or describe the behavior under test.

### Comments

Use comments to explain constraints, decisions, and non-obvious behavior. Do not restate names, types, or control flow, and keep comments current when the code changes.

### Types

Prefer precise types over `any` and broad type assertions. Use narrowing, type guards, or generics when they make the guarantee clearer.

### Naming

Follow the [public API vocabulary](./site/content/docs/custom-controls.mdx#commands-and-queries) when adding or renaming an export.
