# AGENTS.md

Working rules for `@portone-io/docx-editor`, a DOCX editor built on ProseMirror.

## Context routing

Use [README.md](./README.md) for product context and [CONTRIBUTING.md](./CONTRIBUTING.md) for the contribution workflow. Read only the documents relevant to the task.
When adding a reusable guide, add it here with when to read it.
When adding or changing a feature, update the relevant document and revise the surrounding content so its claims remain accurate and ordered by importance.
Keep one source of truth for each concern and link to it instead of repeating its details in another document.

- For an OOXML decision, check [the existing specification notes](./spec/notes/README.md), then follow [the specification guide](./spec/README.md) when no note covers it. Do not load specification prose or schemas for unrelated work.
- Read [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) when adding or replacing redistributed material or its attribution.
- Read [docs/architecture.md](./docs/architecture.md) before moving code, changing folder dependencies, or adding an entry point.
- Read [docs/testing.md](./docs/testing.md) before changing test structure, packaging checks, or the release gate.
- Read [docs/features.md](./docs/features.md) for a support claim or a behavior change.
- Read [docs/core.md](./docs/core.md) for programmatic DOCX import and export.
- Read [docs/custom-controls.md](./docs/custom-controls.md) for commands, queries, plugins, or custom controls.
- Read [site/README.md](./site/README.md) for the documentation and landing site.
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

Follow the [public API vocabulary](./docs/custom-controls.md#commands-and-queries) when adding or renaming an export.
