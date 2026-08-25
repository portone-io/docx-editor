# Testing

Use `pnpm check` for the default local gate. Run the specialized checks when a change affects packaging, a fresh consumer installation, or browser behavior.

## Commands

| Command | Scope |
| --- | --- |
| `pnpm check` | Lint, typecheck, and unit and integration tests |
| `pnpm test` | Vitest tests under `src/` |
| `pnpm typecheck` | TypeScript checks for the package and E2E project |
| `pnpm lint` | Biome checks |
| `pnpm test:package` | Published tarball contents and leaf-import size |
| `pnpm verify:package` | Fresh installation, declarations, entries, bundle, and stylesheet |
| `pnpm test:e2e` | Playwright tests against a locally installed Chrome |

The unit suite requires `xmllint` for OOXML schema validation. `verify:package` also needs network access to install the packed package and its peer dependencies in a temporary project.

## Unit and integration tests

Place a test beside the source it covers, such as `src/docx/importDocx.test.ts` beside `src/docx/importDocx.ts`. Vitest scans `src/` only.

Shared helpers belong under `src/__testing__/` or a feature's `__testing__/` directory. The declaration build excludes those directories, and the package test ensures they are not published.

The suite uses a 30-second timeout because schema validation and tests that exercise compressed-size limits can legitimately take several seconds. [The fixture guide](../__fixtures__/README.md) owns the requirements for committed DOCX files.

## Tests that guard package rules

| Test | Rule it protects |
| --- | --- |
| `src/publicApi.test.ts` | Runtime exports for every JavaScript entry match `api-manifest.json`. |
| `src/folderBoundaries.test.ts` | Folder ranks are respected, every production file is reachable from an entry point, and every production folder is ranked. |
| `src/lockHonesty.test.ts` | A command's applicability result agrees with what it dispatches around locked content. |
| `src/docx/exportSchemaValidation.test.ts` | Every fixture and representative edited export validates against the ECMA-376 Transitional schemas. |

Update `api-manifest.json` only when a public runtime API change is intentional. The lock test lists command factories explicitly so every new command must state how it behaves around locks.

The schema test requires `xmllint`, rejects a missing validator or an empty fixture set, and includes a negative control so a broken validation path cannot pass silently. It removes `mc:Ignorable` before validation as required by the markup-compatibility preprocessing model and supplies the standard XML namespace imported by the schemas.

## Package checks

`pnpm test:package` runs two isolated checks against a clean build:

- `packaging/tarballContents.test.ts` packs the project and verifies exported files, declarations, the documents a consumer reads before installing, excluded development files, and resolved dependency ranges.
- `packaging/leafImportSize.test.ts` rebuilds the output and protects small leaf imports from accidentally pulling in a large shared bundle.

`pnpm verify:package` installs the tarball and its peers in a temporary project outside the repository. It typechecks and bundles a consumer, loads every JavaScript entry, mounts `DocxEditor` over a fixture document in jsdom, and verifies the published stylesheet. This is the check that catches declarations or imports that work only inside the source workspace.

The mount reads `DOCX_EDITOR_REACT_RANGE` to install a different React range in the consumer project, and asserts that a consumer's own ref reaches the handle by exporting the opened document through `downloadDocx`. CI runs the `package` job once per supported React major.

Neither package command is part of `pnpm test`. `verify:package` is kept separate from `pnpm check` because it is slower and needs the network.

## Real-browser tests

`pnpm test:e2e` covers IME composition, focus, keyboard interaction, and layout behavior that jsdom cannot verify. See the [real-browser test guide](../e2e/README.md) for its prerequisites, suite layout, and known limits.

## Workflow audit

The `workflows` job audits the GitHub Actions configuration, `.github/workflows/` and `dependabot.yml` alike, with [zizmor](https://docs.zizmor.sh) under its default `regular` persona. Any finding fails the job and is reported as an annotation on the diff.
Every action is pinned to a commit SHA with the version in a trailing comment, and Dependabot's `github-actions` entry bumps those pins, including the zizmor action that fixes the analyzer version.

Run the same audit locally with `docker run --rm -v "$PWD:/workspace:ro" -w /workspace ghcr.io/zizmorcore/zizmor:1.29.0 .`, matching the version the action pins.

## Release gate

Before a release, run:

```sh
pnpm check
pnpm test:package
pnpm verify:package
DOCX_EDITOR_REACT_RANGE=^18 pnpm verify:package
pnpm test:e2e
```

The development server does not replace `pnpm typecheck`. `pnpm spec 17.5.2.23` is a separate utility for finding an OOXML specification section and is not a test.

Do not weaken or remove a test only to make a change pass. When behavior changes intentionally, update the expectation and explain the decision.
