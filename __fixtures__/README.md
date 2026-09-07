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
| `table-styles.docx` | Export-style package | Table styles that dress the header row, the closing row, the edge columns, a corner and the banded rows |
| `sections-and-revisions.docx` | Export-style package | Two sections with a mid-body section break, a table carrying a grid revision, an inline content control, and a body-level bookmark pair |
| `producers/google-docs-export.docx` | Producer package | Markup Google Docs saved: revision identifiers on every run, a tracked insertion, a generated bookmark name, and measurements the schemas turn down |

A Word-style package includes document properties and named styles, and some also include note parts. An export-style package omits document properties, leaves the default `Normal` style without run properties, and stores list indentation on paragraphs. Keeping both shapes exercises conventions produced by different DOCX writers. A producer package is whatever its producer wrote, which is neither of those shapes on purpose.

## Rules for every fixture

- Build a conformance-lane fixture from package parts; do not save it from a word processor. A producer-lane fixture is made the other way round, and [Producer lane](#producer-lane) is the procedure for it.
- Use A4 paper unless page geometry is the property under test. Give each conformance-lane fixture explicit, nonuniform margins so accidental fallback geometry is visible; a producer-lane file keeps whatever geometry its producer wrote, uniform margins included.
- Do not include real organizations, people, places, addresses, contact details, account numbers, registration numbers, tickets, or authoring metadata. The package name and attributed public-domain text are the only exceptions. In the producer lane `scripts/sanitize-fixture.mjs` is what enforces this, since the producer put a real name in the package.
- Keep one language per file unless multilingual behavior is the purpose of the fixture.
- Use ASCII file names so the live editor can fetch them without additional URL encoding.
- Do not put `mc:AlternateContent` in a conformance-lane fixture. The [validation profile](../docs/testing.md#tests-that-guard-package-rules) selects only one branch, so the others would go untested; write the markup under test directly instead. A producer-lane file keeps whatever its producer wrote, `mc:AlternateContent` included, and there the branch the profile selects is the one under test. `mc:Ignorable` on a part root is expected in either lane; only markup in unknown ignorable namespaces is removed, subject to `mc:ProcessContent`.

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

Three suites read the lane, through `producerFixtureNames` and `readProducerFixture` in [`src/__testing__/docx.ts`](../src/__testing__/docx.ts), and a fourth, `src/sanitizeFixture.test.ts`, reads the same names to hold every committed file to the script under [Sanitizing](#sanitizing):

- `src/docx/producerRoundtrip.test.ts` opens each file, exports it untouched and compares every part, then edits the first paragraph and, separately, a cell of the first table, and checks that the bytes on either side of the edited block are still the producer's own.
- `src/docx/exportSchemaValidation.test.ts` validates the untouched export against an approved list of the violations each producer wrote in, then makes the same two edits and compares each edited export against the untouched one occurrence by occurrence: an edit may add no violation, neither a new kind nor one more of a kind the producer already wrote, and may drop only what [Known gaps](#known-gaps) says a rebuilt block normalizes.
- `src/docx/fidelity.test.ts` records what each file loses on the way in, under `src/docx/__snapshots__/fidelity/producers/`.

### Origin

| File | Producer | Saved |
| --- | --- | --- |
| `google-docs-export.docx` | Google Docs on the web, downloaded as `.docx` | 2026-09-07 |

Google Docs writes no `docProps/` at all, so this file cannot testify to its own origin the way a Word or LibreOffice package does through `docProps/app.xml`. This table is the record instead.

The lane is meant to hold the same document saved from Word, from LibreOffice, and from Google Docs, and so far it holds only the Google Docs file. The lane is therefore not complete: the markup the other two would bring, a deletion suggestion, `mc:AlternateContent` with a floating drawing inside it, a content control, a range bookmark, and a `docProps/` for the sanitizer to act on, is covered by no saved document yet, and [the contract below](#producersgoogle-docs-exportdocx) lists each gap. Adding those two files, through [Sanitizing](#sanitizing) and with a row in this table, is what closes it.

### Sanitizing

```sh
node scripts/sanitize-fixture.mjs downloaded.docx __fixtures__/producers/<name>.docx
```

The script rewrites what names a person, an account, a machine, or a moment, and leaves the rest of the package as the producer wrote it:

- In every part: `w:author` and `w15:author` to `Reviewer A`, `w:initials` to `RA`, `w15:userId` to `reviewer-a`, and `w:date` to `2026-01-01T00:00:00Z`, the day the zip entries are stamped with, since when somebody was editing is authoring metadata as much as who. Only the quote that opened an attribute value closes it, and whitespace around `=` is allowed, as XML allows both.
- In `docProps/core.xml`: `dc:creator` and `cp:lastModifiedBy` to `Fixture Author`, `dcterms:created` and `dcterms:modified` to the same instant, and `dc:title`, `dc:subject`, `dc:description`, `cp:keywords`, and `cp:category` emptied, since they hold whatever the author typed.
- In `docProps/app.xml`: `Company`, `Manager`, and `HyperlinkBase` emptied. `Application` and `AppVersion` stay, because they are how a package names the software that wrote it.
- In `docProps/custom.xml`: every property removed, keeping the part itself, because `_rels/.rels` and `[Content_Types].xml` name it and a package pointing at a part it does not hold is one no producer saved.
- In any `.rels` part: the target of an `attachedTemplate` relationship cut down to its file name, since Word writes the path under the user's profile. The relationship stays, because `w:attachedTemplate` in the settings part points at it.

A part the producer did not write is skipped rather than treated as missing, and a zip holding no `[Content_Types].xml` is refused rather than written out empty. rsids and `w14:paraId` are left alone: they are the markup the lane exists to test.

The list above is what the script has met in a Google Docs package and in the Word-shaped packages `src/sanitizeFixture.test.ts` builds; no Word or LibreOffice download has been through it yet. Before committing the first one, unzip the output and read every part for a name, an account, a path, or a date the list does not cover, and add what turns up to the script and to this list rather than editing the fixture by hand.

Because it repacks with the settings every fixture uses, running it over its own output gives back the same bytes. That equality is the producer-lane replacement for the rebuild-without-edits check above, and it is what to run after replacing a file. `src/sanitizeFixture.test.ts` runs the script over packages shaped like what a producer saves, and over every committed file in the lane, which the script has to give back byte for byte.

### The body text

[`source.md`](./producers/source.md) holds the body every file in the lane carries, so that the same document saved out of different software differs in markup and not in prose.

`google-docs-export.docx` was made by uploading `demo.docx` to Google Docs, leaving a suggestion on it, and downloading the result, so its body arrived reading like the live demo. Google Docs also dropped the demo's content control, turned its endnote into a second footnote, and collapsed its range bookmark to a point, which the list above records. Text that reads like `demo.docx` makes a failing test hard to place, so the body was replaced with `source.md` afterwards, by substituting the contents of `w:t` and the comment body in the saved package. Every element, attribute, rsid, and style reference is still the one Google Docs wrote. A file saved from a producer that this project can drive directly should have `source.md` typed into it instead.

### `producers/google-docs-export.docx`

This fixture must retain:

- `w:rsid*` attributes on runs and paragraphs, and `w14:paraId` on every paragraph of `word/document.xml`, `header1.xml`, `footer1.xml`, `footnotes.xml`, and `comments.xml`.
- One `w:ins` that splits a run mid-word, so revision markup is met inside a paragraph rather than around one.
- A bookmark whose start and end markers stand side by side, so that it is a point and not a range, under a name the producer generated and carrying `w:colFirst` and `w:colLast` on a body paragraph that stands in no table.
- A three-column table with two vertical merges, one `w:gridSpan`, `w:shd` on both cells of each merge, and a `w:tblGridChange` beside its columns; then a second table carrying the three `w:vAlign` values.
- A `w:hyperlink` resolved through an external relationship, an inline `w:drawing` with its media part, two numbering definitions used at two levels each, and a `w:tblStyle` reference on each table.
- Two footnotes, one comment with its range markers, and a complex field written wholly inside one run in `word/footer1.xml`.
- A theme part, and the measurements and the `w:pgMar` that [Known gaps](#known-gaps) accounts for.

What it does not carry, and what a further producer file is wanted for:

- No deletion suggestion. `w:del` and `w:delText` appear nowhere in the package, so no saved document covers them; the Word file planned for this lane is where they come from.
- No bookmark with anything inside it. `demo.docx` carried a range named `FeatureExamples` around a sentence, and Google Docs collapsed it to a point between two sentences and renamed it, so no saved document covers a range marker with content between its ends.
- No content control. Google Docs has none, so the locked `w:sdt` `demo.docx` carries was dropped on the way through.
- No `w:tooltip` on the hyperlink, which Google Docs does not write.
- No endnote part: Google Docs turned the demo's endnote into a second footnote.
- No `mc:AlternateContent`, no floating (`wp:anchor`) drawing, and no `w:lastRenderedPageBreak`.
- No `docProps/` and no `word/people.xml`, so every rewrite the sanitize script makes but the reviewer one has nothing to act on here.

### Known gaps

Every entry in the first table is markup the producer wrote. A block nobody edited goes out as the producer's own bytes without passing through the writer, so these violations are in the export because they were in the file, and no fix is scheduled: correcting them would mean rewriting a block nobody touched, which is the untouched byte identity this lane guards. `PRODUCER_VIOLATIONS` in `src/docx/exportSchemaValidation.test.ts` pins the exact set, one entry per kind of violation, so a producer file that changes, or an export that stops preserving what it was handed, turns that suite red.

| What the schemas turn down | Why |
| --- | --- |
| `w14:paraId` on `w:p`, in five parts | The validation profile discards markup only in namespaces a part declares `mc:Ignorable`, and Google Docs declares none, so the Microsoft extension attribute reaches the validator. The profile in [`src/docx/__testing__/mce.ts`](../src/docx/__testing__/mce.ts) follows the part's own declarations rather than second-guessing them |
| `w:w` on `w:tblW` and on the four `w:tcMar` sides in `word/document.xml`, and on the four `w:tblCellMar` sides of the table styles in `word/styles.xml` | Google Docs writes these measurements with a decimal point, and `ST_MeasurementOrPercent` has no decimal form |
| `w:pgMar` with no `w:gutter` | `CT_PageMar` requires the attribute and Google Docs omits it |

An edited block is different. The writer rebuilds it, writing what it models in the schema's own form and leaving out what it has no model for, so a rebuilt block can stop being turned down for something the producer wrote. `REBUILD_DROPS` in the same suite pins, per file and per kind of edited block, exactly which of the violations above a rebuild makes go away, so that a writer that starts normalizing or dropping more of a producer's markup is approved rather than absorbed. Editing a paragraph of `google-docs-export.docx` drops nothing. Editing a cell of its first table drops three:

| What the rebuilt table no longer carries | Why |
| --- | --- |
| `w:tblW w:w="9026.0"` | The writer writes the table width it models back as the integer `9026` |
| `w14:paraId` on the paragraph of each of the two cells continuing a vertical merge | The writer has no model for the content of a continuation cell and writes it as an empty `w:p`, so the paragraph the producer put there is lost on a rebuild, its identifier and its run included. `blockXml` in [`src/docx/exportDocx.ts`](../src/docx/exportDocx.ts) says so, and the untouched round trip is what keeps such a cell as it was |

The same rebuild writes the producer's `w:vMerge w:val="continue"` as the bare `w:vMerge` the schema defaults to, which is not a violation either way and so appears in neither pin. Whether a rebuilt continuation cell should keep the producer's paragraph is a question for the table writer; this lane records that it does not today.

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

### `table-styles.docx`

This is `size-fallback.docx` with `word/document.xml` and `word/styles.xml` replaced. It is the only fixture whose styles dress the parts of a table separately, so it is what holds the display of a table style. It must retain:

- A table style `ReadingTable` based on the default table style, carrying a `w:tblStylePr` for `wholeTable`, `firstRow`, `lastRow`, `firstCol`, `band1Horz` and `neCell`, each writing something a cell draws: a fill, a line, or bold or italic text.
- `ReadingTableWideBands`, based on it, stating `w:tblStyleRowBandSize` of 2 and nothing else, so that a band of more than one row is covered and so is the layering of a band size down `basedOn`.
- Three tables wearing those styles: one naming the parts it takes through the six `w:tblLook` attributes, one naming them through the legacy `w:val="04A0"` bitmask, and one wearing the wide-band style. The first has four rows, so it has banded rows between its header row and its closing row.
- A `w:cnfStyle` on the first cell of the first table, which the editor works out again from where the cell sits and carries along untouched.
- At least six body paragraphs holding text and a second paragraph style beside the default, which is what the export battery in `src/docx/exportSchemaValidation.test.ts` reserves.

### `sections-and-revisions.docx`

This is `letter-page.docx` with `word/document.xml` replaced. It holds the markup whose identifiers an edit can copy and the export must keep unique. It must retain:

- Two sections on A4: a mid-body paragraph whose `w:pPr` carries a `w:sectPr`, and the body's closing `w:sectPr`, the two differing in every side of `w:pgMar`.
- A 2x2 table whose `w:tblGrid` closes with a `w:tblGridChange`.
- One inline `w:sdt` carrying a `w:id`, and one bookmark pair standing directly under `w:body`, around the table.
- No `w14:paraId`. A copied paragraph's identifier is exercised by unit tests that build a body with `makeDocx`, and a saved document's identifiers are met in the producer lane.

## Live editor

`pnpm dev` opens `demo.docx`. Keep it readable as a product introduction while `kitchen-sink.docx` remains optimized for test coverage.
