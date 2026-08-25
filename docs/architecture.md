# Architecture

This page describes the package boundaries a contributor needs before moving code, adding an entry point, or changing the build.
Its scope is the editor library itself; the other workspace packages, `demo/` and `site/`, document themselves in their own folders.

## Preservation model

Import keeps the original XML behind each document block. An untouched block is written back from that source, and a structure the editor cannot model becomes a placeholder that retains the XML without exposing unsupported edits.

Paragraphs and runs retain their original formatting XML while supported edits replace only the relevant values. Package parts outside the supported editing surface are repacked unchanged.

Import or export fails with a stable error code when the editor cannot guarantee that content will survive the round trip.

## Folder layering

A folder may import itself and folders with a lower rank only.

| Rank | Folder | Responsibility |
| --- | --- | --- |
| 0 | `model` | Shared format values and validation |
| 1 | `styles` | Visual styles, presets, and font fallbacks |
| 2 | `ooxml` | XML, errors, units, and image primitives |
| 3 | `numbering` | Numbering definitions and list markers |
| 4 | `schema` | ProseMirror nodes, marks, rendering, and locks |
| 5 | `docx` | Import, export, session state, and page geometry |
| 6 | `page` | Page-boundary calculations |
| 6 | `table` | Table editing, formatting, and resizing |
| 7 | `editor` | Editor view, plugins, node views, and commands |
| 8 | `ui` | Toolbar, menus, panels, and keyboard behavior |
| 9 | `(root)` | Public entries and the React component |

Folders at the same rank cannot import each other, so `page` and `table` share page data through `docx`.
Subfolders are organizational and inherit the rank of their top-level folder. They split a feature's
parsing, writing, rendering, or interaction responsibilities without creating another layer.
For example, `docx/formatting` separates direct-format parsing from style layering, while
`editor/commands/comments` and `editor/commands/formatting` separate shared models, reads, and edits.

`src/folderBoundaries.test.ts` enforces the ranks, requires every production file to be reachable from an entry point, and rejects an unranked folder. The dependency direction keeps file processing independent from the view layer.

## Entry points

| Subpath | Source | Responsibility |
| --- | --- | --- |
| `.` | `src/index.ts` | React editor, download helper, presets, and extension types |
| `./core` | `src/core.ts` | DOCX import, export, document schema, and package readers |
| `./commands` | `src/editor/commands/index.ts` | Text, paragraph, list, lock, and history operations |
| `./table` | `src/table/index.ts` | Operations on tables already in the document |
| `./styles.css` | `src/styles/editor.css` | Editor styles |

The root entry adds the React editor to the same import and export engine exposed through `./core`. Code reachable from `./core` stays below the editor and UI layers so programmatic document processing does not load a view.

The repository is a pnpm workspace whose root package is the library itself. `demo/` and `site/` are the other two packages: each depends on the library as `workspace:*` and imports `@portone/docx-editor` and `@portone/docx-editor/styles.css`, so both exercise the same entry points a consumer resolves. `demo/` exports the `DocxEditorDemo` component, and `demo/main.tsx` holds the Vite-only shell that loads the fixture and mounts it. The component module stays free of Vite-specific syntax because the site imports it too.

`insertTable` belongs to `./commands` rather than `./table` because a new table uses page geometry stored by the editor layer. Other table commands operate on a table that already exists and do not need that dependency.

## Build and package

`pnpm build` uses esbuild to emit one ESM file per source module and TypeScript to emit declarations beside them. Dependencies and peers remain external, and the stylesheet is copied to `dist/styles.css`.

Development exports point to `src/`, while `publishConfig` maps the same subpaths to `dist/`. `prepack` builds from a clean source tree before packing or publishing.

The one-module-to-one-file ESM output lets compatible bundlers remove unused JavaScript exports. The stylesheet remains listed in `sideEffects` so bundlers do not discard it.

The `files` field limits the package to runtime output and public documentation. `packaging/tarballContents.test.ts` is the authority for what must be present or absent, and [Testing](./testing.md#package-checks) explains the package checks.

## Peer dependencies

ProseMirror packages are peers because objects such as `PluginKey` and `Decoration` depend on module identity. Bundling a second copy would make consumer plugins incompatible with the editor's copy.

A ProseMirror package imported by source code belongs in both `peerDependencies`, for consumers, and `devDependencies`, for this package's tests. The root entry re-exports the common ProseMirror types needed by extension code.

The build leaves all peers external, and the packaging tests ensure published dependency ranges no longer contain workspace catalog syntax.
