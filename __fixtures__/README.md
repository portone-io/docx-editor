# Fixtures

These DOCX packages cover document shapes that the editor must import, edit, preserve, and export. They are built from controlled XML rather than saved from a word processor, so each part is intentional and reproducible.

Tests using `fixtureNames` from [`src/__testing__/docx.ts`](../src/__testing__/docx.ts) automatically run against every file in this directory. A test that needs one particular property names its fixture directly. Tests that need only a small grid invariant build it with [`src/table/__testing__/tables.ts`](../src/table/__testing__/tables.ts) instead.

## Fixture map

| File | Shape | Purpose |
| --- | --- | --- |
| `kitchen-sink.docx` | Word-style package | Broad coverage of formatting, lists, tables, controls, links, images, and preserved parts |
| `demo.docx` | Word-style package | Readable product demonstration opened by the live editor |
| `size-fallback.docx` | Export-style package | Font-size fallbacks, indented lists, narrow and oversized tables, and multiple pages |
| `east-asian.docx` | Word-style package | Per-script fonts, theme fonts, language metadata, and CJK line-breaking properties |
| `letter-page.docx` | Export-style package | US Letter geometry and margins |

A Word-style package includes document properties and named styles, and some also include note parts. An export-style package omits document properties, leaves the default `Normal` style without run properties, and stores list indentation on paragraphs. Keeping both shapes exercises conventions produced by different DOCX writers.

## Rules for every fixture

- Build fixtures from package parts; do not save them from a word processor.
- Use A4 paper unless page geometry is the property under test. Give each fixture explicit, nonuniform margins so accidental fallback geometry is visible.
- Do not include real organizations, people, places, addresses, contact details, account numbers, registration numbers, tickets, or authoring metadata. The package name and attributed public-domain text are the only exceptions.
- Keep one language per file unless multilingual behavior is the purpose of the fixture.
- Use ASCII file names so the live editor can fetch them without additional URL encoding.

Prefer original text. If a fixture uses redistributable text from elsewhere, record its source and status here:

| Author | Work | Status | Files |
| --- | --- | --- | --- |
| Lewis Carroll (1832–1898) | *Alice's Adventures in Wonderland* (1865), *Through the Looking-Glass* (1871) | Public domain | `size-fallback.docx`, `letter-page.docx` |

## Adding or replacing a fixture

Unzip an existing fixture, edit its package parts, and repack it with fflate as `zipSync(parts, { level: 6, mtime: new Date(2026, 0, 1) })`. First rebuild the file without edits and verify that its bytes are identical; then apply the intended change through the same process.

Run `pnpm test` after adding or replacing a file. Every WordprocessingML part must validate against `spec/schemas/` before and after export. If a replacement breaks a test, restore the property in the fixture rather than weakening the assertion.

Add the file to the fixture map, document any properties that direct tests depend on, and add an attribution row when its text comes from another source.

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
