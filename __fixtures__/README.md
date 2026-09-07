# Fixtures

These DOCX packages cover document shapes that the editor must import, edit, preserve, and export. They come in two lanes.

The files in this directory are the conformance lane: built from controlled XML rather than saved from a word processor, so each part is intentional and reproducible. The files under `producers/` are the [producer lane](#producer-lane): saved by a word processor and then sanitized, so that the corpus also holds markup nobody would write on purpose.

Tests using `fixtureNames` from [`src/__testing__/docx.ts`](../src/__testing__/docx.ts) automatically run against every file in this directory. `fixtureNames` reads only the top level, so a producer file is never picked up by a suite that did not ask for it; `producerFixtureNames` beside it is how a suite opts into that lane. A test that needs one particular property names its fixture directly. Tests that need only a small grid invariant build it with [`src/table/__testing__/tables.ts`](../src/table/__testing__/tables.ts) instead.

## Fixture map

| File | Shape | Purpose |
| --- | --- | --- |
| `kitchen-sink.docx` | Word-style package | Broad coverage of formatting, lists, tables, controls, links, images, and preserved parts |
| `demo.docx` | Word-style package | Readable product demonstration opened by the live editor |
| `size-fallback.docx` | Export-style package | Font-size fallbacks, indented lists, narrow and oversized tables, and multiple pages |
| `east-asian.docx` | Word-style package | Per-script fonts, theme fonts, language metadata, and CJK line-breaking properties |
| `letter-page.docx` | Export-style package | US Letter geometry and margins |
| `producers/google-docs-export.docx` | Producer package | Markup Google Docs saved: revision identifiers on every run, a tracked insertion, a generated bookmark name, and measurements the schemas turn down |

A Word-style package includes document properties and named styles, and some also include note parts. An export-style package omits document properties, leaves the default `Normal` style without run properties, and stores list indentation on paragraphs. Keeping both shapes exercises conventions produced by different DOCX writers. A producer package is whatever its producer wrote, which is neither of those shapes on purpose.

## Rules for every fixture

- Build a conformance-lane fixture from package parts; do not save it from a word processor. A producer-lane fixture is made the other way round, and [Producer lane](#producer-lane) is the procedure for it.
- Use A4 paper unless page geometry is the property under test. Give each conformance-lane fixture explicit, nonuniform margins so accidental fallback geometry is visible; a producer-lane file keeps whatever geometry its producer wrote, uniform margins included.
- Do not include real organizations, people, places, addresses, contact details, account numbers, registration numbers, tickets, or authoring metadata. The package name and attributed public-domain text are the only exceptions. In the producer lane `scripts/sanitize-fixture.mjs` is what enforces this, since the producer put a real name in the package.
- Keep one language per file unless multilingual behavior is the purpose of the fixture.
- Use ASCII file names so the live editor can fetch them without additional URL encoding.
- Do not put `mc:AlternateContent` in a fixture. The [validation profile](../docs/testing.md#tests-that-guard-package-rules) selects only one branch, so other branches would not be tested. Write the markup under test directly instead. `mc:Ignorable` on a part root is expected; only markup in unknown ignorable namespaces is removed, subject to `mc:ProcessContent`.

Prefer original text. If a fixture uses redistributable text from elsewhere, record its source and status here:

| Author | Work | Status | Files |
| --- | --- | --- | --- |
| Lewis Carroll (1832–1898) | *Alice's Adventures in Wonderland* (1865), *Through the Looking-Glass* (1871) | Public domain | `size-fallback.docx`, `letter-page.docx` |

## Adding or replacing a fixture

Unzip an existing fixture, edit its package parts, and repack it with fflate as `zipSync(parts, { level: 6, mtime: new Date(2026, 0, 1) })`. First rebuild the file without edits and verify that its bytes are identical; then apply the intended change through the same process. A producer-lane file has no earlier version in the repository to rebuild, so [Producer lane](#producer-lane) names the equivalent check.

Run `pnpm test` after adding or replacing a file. Every WordprocessingML part must validate against `spec/schemas/` before and after export, bar the violations a producer wrote in and [Known gaps](#known-gaps) accounts for. If a replacement breaks a test, restore the property in the fixture rather than weakening the assertion.

Add the file to the fixture map, document any properties that direct tests depend on, and add an attribution row when its text comes from another source.

## Producer lane

`producers/` holds documents a word processor saved. Their markup is the one thing this repository cannot invent: a producer stamps `w:rsid*` identifiers on every run, `w14:paraId` on every paragraph, writes its own defaults over properties no author ever set, and puts values in the package that the published schemas turn down. Writing that by hand would be committing a guess about what real software does.

Three suites read the lane, through `producerFixtureNames` and `readProducerFixture` in [`src/__testing__/docx.ts`](../src/__testing__/docx.ts):

- `src/docx/producerRoundtrip.test.ts` opens each file, exports it untouched and compares every part, then edits the first editable paragraph and checks that the bytes on either side of it are still the producer's own.
- `src/docx/exportSchemaValidation.test.ts` validates the untouched and the edited export, against an approved list of the violations each producer wrote in.
- `src/docx/fidelity.test.ts` records what each file loses on the way in, under `src/docx/__snapshots__/fidelity/producers/`.

### Origin

| File | Producer | Saved |
| --- | --- | --- |
| `google-docs-export.docx` | Google Docs on the web, downloaded as `.docx` | 2026-09-07 |

Google Docs writes no `docProps/` at all, so this file cannot testify to its own origin the way a Word or LibreOffice package does through `docProps/app.xml`. This table is the record instead.

### Sanitizing

```sh
node scripts/sanitize-fixture.mjs downloaded.docx __fixtures__/producers/<name>.docx
```

The script sets `dc:creator` and `cp:lastModifiedBy` to `Fixture Author`, empties `Company`, deletes `docProps/custom.xml`, and rewrites every `w:author`, `w:initials`, and `w15:author` in every part to `Reviewer A` and `RA`. A part the producer did not write is skipped rather than treated as missing.

It leaves rsids, `w14:paraId`, and `docProps/app.xml`'s `Application` and `AppVersion` alone: the first two are the markup the lane exists to test, and the last is how a package names the software that wrote it.

Because it repacks with the settings every fixture uses, running it over its own output gives back the same bytes. That equality is the producer-lane replacement for the rebuild-without-edits check above, and it is what to run after replacing a file.

### The body text

[`source.md`](./producers/source.md) holds the body every file in the lane carries, so that the same document saved out of different software differs in markup and not in prose.

`google-docs-export.docx` was made by uploading `demo.docx` to Google Docs, leaving a suggestion on it, and downloading the result, so its body arrived reading like the live demo. Text that reads like `demo.docx` makes a failing test hard to place, so the body was replaced with `source.md` afterwards, by substituting the contents of `w:t` and the comment body in the saved package. Every element, attribute, rsid, and style reference is still the one Google Docs wrote. A file saved from a producer that this project can drive directly should have `source.md` typed into it instead.

### `producers/google-docs-export.docx`

This fixture must retain:

- `w:rsid*` attributes on runs and paragraphs, and `w14:paraId` on every paragraph of `word/document.xml`, `header1.xml`, `footer1.xml`, `footnotes.xml`, and `comments.xml`.
- One `w:ins` that splits a run mid-word, so revision markup is met inside a paragraph rather than around one.
- A bookmark pair whose name the producer generated, carrying `w:colFirst` and `w:colLast` on a body paragraph that stands in no table.
- A three-column table with two vertical merges, one `w:gridSpan`, `w:shd` on both cells of each merge, and a `w:tblGridChange` beside its columns; then a second table carrying the three `w:vAlign` values.
- A `w:hyperlink` resolved through an external relationship, an inline `w:drawing` with its media part, two numbering definitions used at two levels each, and a `w:tblStyle` reference on each table.
- Two footnotes, one comment with its range markers, and a complex field written wholly inside one run in `word/footer1.xml`.
- A theme part, and the measurements and the `w:pgMar` that [Known gaps](#known-gaps) accounts for.

What it does not carry, and what a further producer file is wanted for:

- No deletion suggestion. `w:del` and `w:delText` appear nowhere in the package, so no saved document covers them; the Word file planned for this lane is where they come from.
- No content control. Google Docs has none, so the locked `w:sdt` `demo.docx` carries was dropped on the way through.
- No `w:tooltip` on the hyperlink, which Google Docs does not write.
- No endnote part: Google Docs turned the demo's endnote into a second footnote.
- No `mc:AlternateContent`, no floating (`wp:anchor`) drawing, and no `w:lastRenderedPageBreak`.
- No `docProps/` and no `word/people.xml`, so every rewrite the sanitize script makes but the reviewer one has nothing to act on here.

### Known gaps

Every entry below is markup the producer wrote and this editor hands back untouched, so it is the producer's output that the schemas turn down rather than the export's. `PRODUCER_VIOLATIONS` in `src/docx/exportSchemaValidation.test.ts` pins the exact set, one entry per kind of violation, so closing a gap turns that suite red and the fix is approved rather than absorbed.

| What the schemas turn down | Why | Whose fix |
| --- | --- | --- |
| `w14:paraId` on `w:p`, in five parts | The validation profile discards markup only in namespaces a part declares `mc:Ignorable`, and Google Docs declares none, so the Microsoft extension attribute reaches the validator | The MCE profile in [`src/docx/__testing__/mce.ts`](../src/docx/__testing__/mce.ts), which has to decide whether a profile drops the extension namespaces regardless of what a part declares |
| `w:w` on `w:tblW` and on the four `w:tcMar` sides, in `word/document.xml` and `word/styles.xml` | Google Docs writes these measurements with a decimal point, and `ST_MeasurementOrPercent` has no decimal form | The XML writing layer, which owns whether a preserved measurement may be normalized on the way out. It may not today, because normalizing it would break the untouched byte identity this lane guards |
| `w:pgMar` with no `w:gutter` | `CT_PageMar` requires the attribute and Google Docs omits it | The same, for the same reason |

## Fixture contracts

Record durable, non-obvious properties shared across tests or required for interoperability. Keep one-off copy, spacing, and visual ordering in focused tests instead of listing them here.

### `kitchen-sink.docx`

This is the broad test fixture. It must retain:

- Direct formatting, paragraph styles, alignment, indentation, line spacing, and line and page breaks.
- Multilevel numbered and bulleted lists, including the marker sequence `1.`, `1)`, `2)`, `3)`, `4)` and a table-cell list using its own definition.
- Two tables in order: first a grid with vertical and horizontal merges, then a table whose required `w:tblGrid` contains no `w:gridCol` elements.
- Percentage table width, runs with and without explicit sizes, and widths used by table serialization tests.
- Locked inline and table-cell content controls.
- Header, footer, footnote, and endnote parts.
- External, bookmark, and unresolved hyperlink paths, including a preserved `w:tooltip`.
- An inline image with its media part, relationship, and content type.

It is allowed to look like a test document rather than a product demo.

### `demo.docx`

This fixture is a readable capability overview for the live editor. The shared round-trip, validation, and live-editor layout suites exercise it.

Keep its claims consistent with the document itself. Detailed support claims belong in [Feature support](../site/content/docs/features.mdx), not in the fixture.

### `size-fallback.docx`

This fixture must retain:

- No `w:sz` in document defaults, the `Normal` style, or table runs, and no paragraph style references.
- Latin and Korean font names that exercise the fallback groups in [`src/styles/fontStack.ts`](../src/styles/fontStack.ts).
- Paragraph-level list indentation, the narrowest committed column, and a table wider than its body.
- Enough content to reach a second page, plus a header and footer.

### `east-asian.docx`

This fixture must retain:

- Original Japanese, Chinese, and Korean body text with English titles, headings, and table labels.
- Distinct `ascii`, `hAnsi`, and `eastAsia` font slots, including mixed Latin and Japanese text in one run.
- Major and minor theme-font references with a custom theme font scheme.
- East Asian `w:lang` values and the `w:kinsoku`, `w:wordWrap`, `w:overflowPunct`, `w:autoSpaceDE`, `w:autoSpaceDN`, and `w:eastAsianLayout` properties.
- Explicit sizes, no numbered paragraphs, a numbering part available for new lists, and one unmerged table.

### `letter-page.docx`

This is `size-fallback.docx` with only `w:pgSz` and `w:pgMar` changed. Keep it as a committed US Letter document so geometry tests do not validate a reader against values produced by the same code under test.

## Live editor

`pnpm dev` opens `demo.docx`. Keep it readable as a product introduction while `kitchen-sink.docx` remains optimized for test coverage.
