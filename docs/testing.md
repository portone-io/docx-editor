# Testing

Use `pnpm check` for the default local gate. Run the specialized checks when a change affects packaging, a fresh consumer installation, or browser behavior.

## Commands

| Command | Scope |
| --- | --- |
| `pnpm check` | Lint, typecheck, and unit and integration tests, on React 19 |
| `pnpm test` | Vitest tests under `src/` |
| `pnpm typecheck` | TypeScript checks for the package and E2E project |
| `pnpm lint` | Biome checks |
| `pnpm check:react18` | The same typecheck and tests, resolved against React 18 |
| `pnpm test:react18` | Vitest tests under `src/` on React 18 |
| `pnpm typecheck:react18` | TypeScript checks for `src/` against the React 18 declarations |
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

## React 18

The package supports React 18 and React 19, so the checks that touch React run twice.
The repository develops on React 19, and the React 18 path resolves the same source and the same tests against React 18 instead.

```sh
pnpm check:react18
```

`react18/` is a workspace package that holds nothing but the React 18 installation.
It exists because pnpm satisfies `react-dom`'s `react` peer from the importer: an `npm:react@18` alias in the root manifest would still be paired with the React 19 that the rest of the repository installs, and the two copies refuse to run together.
A separate importer gets the pairing right.

[`vitest.react18.config.ts`](../vitest.react18.config.ts) points the React entries at that installation, and so does every runtime dependency that carries a React copy of its own, because React 18 cannot render elements React 19 built.
[`tsconfig.react18.json`](../tsconfig.react18.json) does the same for the declarations.
A dependency with a `react` peer therefore has to be added to `react18/package.json` and to the alias list, or the suite fails on the version it was meant to prove.

`pnpm verify:package` reads `DOCX_EDITOR_REACT_RANGE` and installs that range in the consumer project instead of the one the manifest declares, which is how the published package is proven against each supported major:

```sh
DOCX_EDITOR_REACT_RANGE=^18 pnpm verify:package
```

CI runs the `check` and `package` jobs once per supported React major.

## Package checks

`pnpm test:package` runs two isolated checks against a clean build:

- `packaging/tarballContents.test.ts` packs the project and verifies exported files, declarations, the documents a consumer reads before installing, excluded development files, and resolved dependency ranges.
- `packaging/leafImportSize.test.ts` rebuilds the output and protects small leaf imports from accidentally pulling in a large shared bundle.

`pnpm verify:package` installs the tarball and its peers in a temporary project outside the repository. It typechecks and bundles a consumer, loads every JavaScript entry, mounts `DocxEditor` over a fixture document in jsdom, and verifies the published stylesheet. This is the check that catches declarations or imports that work only inside the source workspace.

The mount is what proves the published build against a React the repository does not develop on. It renders the editor through a consumer's own ref, exports the opened document through the handle, and requires `downloadDocx` to report `exported` rather than `unavailable`, so a ref that stops reaching the handle is caught outside the workspace as well as inside it.

Neither package command is part of `pnpm test`. `verify:package` is kept separate from `pnpm check` because it is slower and needs the network.

## Real-browser tests

`pnpm test:e2e` covers IME composition, focus, keyboard interaction, and layout behavior that jsdom cannot verify. See the [real-browser test guide](../e2e/README.md) for its prerequisites, suite layout, and known limits.

## Release gate

Before a release, run:

```sh
pnpm check
pnpm check:react18
pnpm test:package
pnpm verify:package
DOCX_EDITOR_REACT_RANGE=^18 pnpm verify:package
pnpm test:e2e
```

The development server does not replace `pnpm typecheck`. `pnpm spec 17.5.2.23` is a separate utility for finding an OOXML specification section and is not a test.

Do not weaken or remove a test only to make a change pass. When behavior changes intentionally, update the expectation and explain the decision.
