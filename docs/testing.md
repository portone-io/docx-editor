# Testing

Use `pnpm check` for the default local gate. Run the specialized checks when a change affects packaging, a fresh consumer installation, or browser behavior.

## Commands

| Command | Scope |
| --- | --- |
| `pnpm check` | Lint, typecheck, unit and integration tests, site release tests, and demo dependency agreement |
| `pnpm test` | Vitest tests under `src/` |
| `pnpm typecheck` | TypeScript checks for the package and E2E project |
| `pnpm lint` | Biome checks |
| `pnpm test:package` | Published tarball contents, leaf-import size, declaration reports, and the core entry in a Node runtime with no DOM |
| `pnpm verify:package` | Fresh installation, declarations, entries, bundle, and stylesheet |
| `pnpm test:e2e` | Playwright tests against a locally installed Chrome |
| `pnpm check:demo-library` | The site and demo pins match their installed packages and do not resolve to the working tree |
| `pnpm test:site-release` | Offline Node tests for released-demo preparation and the site version commit |

The unit suite requires `xmllint` for OOXML schema validation. `verify:package` also needs network access to install the packed package and its peer dependencies in a temporary project.

`check:demo-library` never reads the registry. `test:site-release` runs `scripts/*.test.mjs` using Node's test runner and temporary files, with registry, installation, and GitHub requests substituted. It also runs the actual Changesets version command in a temporary workspace to ensure release preparation preserves published demo pins. It covers release visibility delays, bounded retries, exact-version selection, installation and build failures, rollback of input files, downgrade prevention, and atomic commits when `main` changes. It neither publishes packages nor pushes commits. [The site guide](../site/README.md#the-version-the-demo-runs) explains the live update path.

`pnpm spec 17.5.2.23` looks up an OOXML specification section. It is a utility, not a test.

## Unit and integration tests

Place a test beside the source it covers, such as `src/docx/importDocx.test.ts` beside `src/docx/importDocx.ts`. Vitest scans `src/` only.

Shared helpers belong under `src/__testing__/` or a feature's `__testing__/` directory. The declaration build excludes those directories, and the package test ensures they are not published.

The suite uses a 30-second timeout because schema validation and tests that exercise compressed-size limits can legitimately take several seconds. [The fixture guide](../__fixtures__/README.md) owns the requirements for committed DOCX files.

## Tests that guard package rules

| Test | Rule it protects |
| --- | --- |
| `src/publicApi.test.ts` | Runtime exports for every JavaScript entry match `api-manifest.json`. |
| `src/folderBoundaries.test.ts` | Folder ranks are respected, every production file is reachable from an entry point, every production folder is ranked, and no two modules read each other at runtime. |
| `src/lockHonesty.test.ts` | A command's applicability result agrees with what it dispatches around locked content, across the bookmark and note markers a document is preserved with, and under every editing protection. Each place also states which guards refuse there, and the stated guards are held against the ones that answer. |
| `src/docx/exportSchemaValidation.test.ts` | Every fixture and representative edited export validates against the ECMA-376 Transitional schemas. |
| `src/docx/writerProbes.test.ts` | Every export of `./commands` and `./table` either runs before a validated export as a writer probe or states why it reaches no writer. |
| `src/schema/domRoundtrip.test.ts` | Every fixture survives being drawn to the DOM and read back, which is the path an IME composition takes. |
| `src/schema/rawAttrs.test.ts` | Every attr the writer writes from says whether it carries raw XML, and each one that does is drawn holding its shape and not holding it. |
| `packaging/apiReport.test.ts` | The committed `etc/*.api.md` reports match the declarations built from each published entry point. |
| `packaging/coreRuntime.test.ts` | The built `dist/core.js` opens every fixture in a Node runtime holding no DOM globals, reading through the `xmlParser` option and reading the export back, answers the returned-file verifier there, and refuses with `no-xml-parser` when given neither the option nor a `DOMParser` global. |

Update `api-manifest.json` only when a public runtime API change is intentional. `pnpm api:update` does the same for the declaration reports, which record types and signatures rather than names. The lock test lists command factories explicitly so every new command must state how it behaves around locks and markers and under every editing protection. The probe test reads the same manifest, so a new command must also say what it writes into an exported package.

Each writer probe has a required `check(before, after)` for its immediate effect. The battery rejects display-only changes and exports every intermediate result before the next command can overwrite it. Every result must change an exported part, every XML part is parsed, and all distinct WordprocessingML outputs are validated in one batch. Optional package-wide assertions run on the final result as well. A setup step belongs in `prepare`, so the probe is measured against the state immediately before its own command.

The schema test requires `xmllint`, rejects a missing validator or an empty fixture set, and includes negative controls so a broken validation path cannot pass silently. Its MCE profile follows ECMA-376 Part 3 sections 7 and 9 for `Ignorable`, `ProcessContent`, `MustUnderstand`, and alternate content. The understood namespaces come from the imports of the committed WML schema and the supplied XML namespace schema. Known namespaces remain subject to validation even if declared ignorable. For unknown ignorable elements, `ProcessContent` preserves their children for validation; otherwise the subtree is removed. `AlternateContent` selects the first Choice whose required namespaces are understood, or its Fallback if no Choice matches. Declarations are resolved in their original scope before wrappers are removed.

This is a validation profile, not a complete MCE consumer: preservation hints and application-defined extension-element configurations are unsupported. Unsupported directives, unbound prefixes, and malformed alternate-content branch structure fail explicitly. Hand-built fixtures must not use `AlternateContent` to hide the markup they are meant to exercise; [the fixture guide](../__fixtures__/README.md) records that rule.

Parts with no committed validation schema, including relationships, content types, people, and commentsExtended, are checked for well-formedness only. Parsing does not verify their vocabulary or cross-part references.

## Package checks

`pnpm test:package` runs four isolated checks against a clean build:

- `packaging/tarballContents.test.ts` packs the project and verifies exported files, declarations, the documents a consumer reads before installing, excluded development files, and resolved dependency ranges.
- `packaging/leafImportSize.test.ts` rebuilds the output and protects small leaf imports from accidentally pulling in a large shared bundle.
- `packaging/apiReport.test.ts` rebuilds the declarations and fails when a committed report in `etc/` no longer matches them.
- `packaging/coreRuntime.test.ts` loads the built core entry in a Node environment with no DOM globals, which is the runtime a server verifying a returned file actually has. The rest of the suite runs under jsdom, so this is the only place a global reached for by accident shows up as a failure.

`pnpm verify:package` installs the tarball and its peers in a temporary project outside the repository. It typechecks and bundles a consumer, loads every JavaScript entry, mounts `DocxEditor` over a fixture document in jsdom, and verifies the published stylesheet. It packs that tarball itself unless `DOCX_EDITOR_TARBALL` points it at one. This is the check that catches declarations or imports that work only inside the source workspace.

The mount reads `DOCX_EDITOR_REACT_RANGE` to install a different React range in the consumer project, and asserts that a consumer's own ref reaches the handle by exporting the opened document through `downloadDocx`. The `package` job verifies every supported React major on one runner, in parallel, over a tarball it packs once, so the suite reads as a single check.

Neither package command is part of `pnpm test`. `verify:package` is kept separate from `pnpm check` because it is slower and needs the network.

## Real-browser tests

`pnpm test:e2e` covers IME composition, focus, keyboard interaction, and layout behavior that jsdom cannot verify. See the [real-browser test guide](../e2e/README.md) for its prerequisites, suite layout, and known limits.

## Workflow audit

The `workflows` job audits the GitHub Actions configuration, `.github/workflows/` and `dependabot.yml` alike, with [zizmor](https://docs.zizmor.sh) under its default `regular` persona. Any finding fails the job and is reported as an annotation on the diff.
Every action is pinned to a commit SHA with the version in a trailing comment, and Dependabot's `github-actions` entry bumps those pins, including the zizmor action that fixes the analyzer version.

Run the same audit locally with `docker run --rm -v "$PWD:/workspace:ro" -w /workspace ghcr.io/zizmorcore/zizmor:1.29.0 .`, matching the version the action pins.

## Keeping the suite honest

Do not weaken or remove a test only to make a change pass.
When behavior changes intentionally, update the expectation and explain the decision.

Every check above runs on a pull request, and [releasing](./releasing.md) re-runs the package ones before anything reaches npm.
