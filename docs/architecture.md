# Architecture

This page describes the package boundaries a contributor needs before moving code, adding an entry point, or changing the build.
Its scope is the editor library itself; the other workspace packages, `demo/` and `site/`, document themselves in their own folders.

## Preservation model

Import keeps the original XML behind each document block. An untouched block is written back from that source, and a structure the editor cannot model becomes a placeholder that retains the XML without exposing unsupported edits. There is one placeholder node kind, `rawBlock`, wherever such a block stands: one opened under the body names the session fragment it was read from, one opened inside a table cell carries its XML, and either may be moved into the other's place by an edit.

Each side story - a comment's body, a footnote's - is read by the same block readers as the main story and kept twice: as it arrived on the session, and as it currently stands in `doc.attrs.stories` under the key naming it (`docx/story`). That is what makes an edit to one an ordinary transaction, undoable and comparable, and it is why an untouched block of a comment body goes back out as its own bytes just as a body block does.

The final body section is held in `doc.attrs.sectPr` and written after the blocks. A body containing only section properties gets an empty editable paragraph whose original XML is empty: export omits it while it remains unchanged and alone, preserving the untouched document. Once another block is added, the paragraph is written as a real blank line.

Paragraphs and runs retain their original formatting XML while supported edits replace only the relevant values. Package parts outside the supported editing surface are repacked unchanged.

The inline wrappers a stretch of text stands inside - a content control, a hyperlink - are a mark each in one group, carrying the depth the file nested them at, and `docx/wrappers` holds what each kind reads out of its element and writes back around the same content. The paragraph reader recurses into them and the paragraph writer groups from the outermost inwards, and the import policy takes its wrapper rows from the registry, so none of the three knows one kind from another. What a further kind costs is a mark spec carrying `WRAPPER_ATTRS`, a registry entry, and a row in `schema/attrRoles` for the attrs it adds; a mark that excludes its own type records one of itself at a time, and the inner of two such wrappers is kept whole where it stood.

Raw XML reaches the model from three directions: import puts what it cut out of the file straight into the attrs, a command builds a fragment of its own (`editor/commands/lockCommands` writes a control's opening tag), and the page is read back. The last of those is the only one that carries a string from outside, so every raw attr the schema's `parseDOM` rules read goes through `ooxml/fragment`, which holds a fragment against the shape the attr carrying it goes back out as. A rule that meets a refusal gives up, so the content settles one level plainer instead of carrying a fragment the writer would splice into the exported file.
What arrives on the clipboard is not one of those directions: `editor/clipboard` reads it with readers of its own, so a paste never reaches a `parseDOM` rule and a copy carries no raw XML for one to read.
Image handlers honor the text-only mode selected by the clipboard pipeline. An empty reading falls back to clipboard text or leaves the selection intact; every nonempty slice reaches ProseMirror's insertion or the table's cell-selection handler.

The fragment gate checks every explicit namespace declaration in the parsed subtree, so a nested declaration cannot hide a rebinding of `w` or `r`. A fragment whose namespace declaration stayed on the original part is checked by local name; the wrapper's placeholder namespace is not evidence of a foreign namespace. Explicit foreign bindings remain rejected for named element shapes.

Each node and mark attr declares a role in `schema/attrRoles.ts`: `source` is what the writer writes from, `display` is worked out from the source and the formatting around it and is never written into a block's XML, and `session` is a value export reads but neither writes nor re-derives, such as which block a node was opened from or whether a comment came in with the file.
Export compares two nodes with `sameSource`, which reads the source and session attrs and ignores the display ones.
The distinction matters because opening a document works its display values out again - a table's shared cell borders among them - and a block judged changed is a block rebuilt, which costs it the markup the writer does not model.

### Display derivation

A display value is worked out again whenever what it depends on moves, through one walk over the document's blocks (`schema/displayDerivation`).
A deriver answers for the node types it names and hands back, for one node, the display attrs that node, its descendants, and their existing run marks should carry; the walk writes one step per node whose attrs would change and reads only the display attrs of what a deriver hands back, so a deriver cannot write a source attr.
The interface stands in `schema`, below `docx` and `editor`, so that a module of either can write a deriver; `editor/plugins/displayDerivation` registers them (`paragraphDisplay` for paragraph and run style values, `tableDisplay` for the lines of a table's cells) and runs the walk after every edit, over the nodes the edit moved, and over every node when the document snapshot (`editor/editorDocument`) is replaced.
A state is built over a document whose every display value was worked out against the state's own snapshot, so a document opened without one draws as one that laid nothing down.
A note reference's label is a display value of an inline node, which the walk does not reach: `editor/plugins/noteNumbering` works the labels out again through `docx/notes/numbering` after a change that moves a reference or a section break, under the same pass and the same history rule, and import labels the opened document through the same function.

The transaction the walk appends carries the display-only pass: `transactionAllowed` in `schema/guards` lets a transaction through every guard when it carries the pass and every step of it, judged off the role table alone, changes display attrs and nothing else.
The pass is a claim rather than a key: a step that rewrites a source attr, a lock flag included, or puts content anywhere fails it, and the transaction is judged as any edit.
Replacing an existing run mark is also display-only when its source and session attrs remain identical throughout the range. New marks and changed XML remain edits.

The re-derivation after an edit goes to the history with that edit; the one after a snapshot change goes to the history not at all, since the values follow the snapshot.

Export writes every part beside the body through a list of part planners (`docx/partPlan`).
A planner answers with the parts it rewrites and declares what a part it adds needs through the one relationships writer and the one content types writer every planner shares, so those two parts are written once from everything asked for; a part's children are put in by `ooxml/partSplice`, which cuts the original text at its root and leaves every other byte as it arrived.
The list is `docx/partPlanners`, where each planner stands beside the side stories it writes and the changes to one it carries; the writer runs the planners off that list and the export invariants read what is written off the same list, so a change no planner carries is refused rather than dropped.
Whether a side story changed is judged once, in `docx/storyParts`, for every story writer and the export invariants alike; a part holding one entry per story, such as the Footnotes part, is written by `storyEntriesPlanner`.
Every rewritten part is parsed before the package is repacked.

A protection level is a policy object in `docx/protectionPolicy`: the package parts it lets a change rewrite, the grammar their entries are written in, and how the document story is compared once its own markup is taken out.
The part planners and the server verifier read the same object, so a part the writer starts adding is a part the verifier already excuses from the byte comparison and judges entry by entry.

Import or export fails with a stable error code when the editor cannot guarantee that content will survive the round trip.

## Folder layering

A folder may import itself and folders with a lower rank only.

| Rank | Folder | Responsibility |
| --- | --- | --- |
| 0 | `model` | Shared format values and validation |
| 1 | `styles` | Visual styles, presets, and font fallbacks |
| 2 | `ooxml` | XML reading and writing primitives, child order, part splicing, errors, units, and image primitives |
| 3 | `numbering` | Numbering definitions and list markers |
| 4 | `schema` | ProseMirror nodes, marks, rendering, locks, and edit guards |
| 5 | `docx` | Import, export, session state, and page geometry |
| 6 | `page` | Page-boundary calculations; `page/kinds` holds one measurer and decorator per breakable block shape |
| 6 | `table` | Table editing, formatting, and resizing |
| 7 | `editor` | Editor view, the document snapshot, plugins, node views, and commands |
| 8 | `ui` | Toolbar, menus, panels, and keyboard behavior |
| 9 | `(root)` | Public entries and the React component |

Folders at the same rank cannot import each other, so `page` and `table` share page data through `docx`.
Subfolders are organizational and inherit the rank of their top-level folder. They split a feature's
parsing, writing, rendering, or interaction responsibilities without creating another layer.
For example, `docx/formatting` separates direct-format parsing, the run property table, style layering, and the hierarchy resolver, `docx/notes` reads the footnotes and endnotes parts, numbers the references to them, and writes the footnotes part, `page/kinds`
gives each breakable block shape its own measurer and decorator, `page/demands` holds what asks a page for room at its foot, `editor/stories` draws a side story, `ui/notes` places footnotes and endnotes around the page, `editor/clipboard` holds the clipboard's props and the readers a paste is read with
in one plugin and leaves what a copy goes out as to `schema/clipboard`, while
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

The repository is a pnpm workspace whose root package is the library itself. `demo/` and `site/` are the other two packages: each imports `@portone/docx-editor` and `@portone/docx-editor/styles.css`, so both exercise the same entry points a consumer resolves. `demo/` exports the `DocxEditorDemo` component, and `demo/main.tsx` holds the Vite-only shell that loads the fixture and mounts it. The component module stays free of Vite-specific syntax because the site imports it too.

Both pin the library to an exact released version rather than depending on it as `workspace:*`, because the site's landing page demonstrates a version to a visitor and names it on a badge. `demo/vite.config.ts` aliases those entry points back to `src/` so `pnpm dev` and `pnpm build:demo` still read the working tree. [The site guide](../site/README.md#the-version-the-demo-runs) owns automatic release updates, local site preparation, and the offline check that guards the installed version.

`insertTable` belongs to `./commands` rather than `./table` because a new table uses page geometry stored by the editor layer. Other table commands operate on a table that already exists and do not need that dependency.

## Build and package

`pnpm build` uses esbuild to emit one ESM file per source module and TypeScript to emit declarations beside them. Dependencies and peers remain external, and the stylesheet is copied to `dist/styles.css`. A committed report under `etc/` for each entry point records that declared surface, and `pnpm api:update` regenerates it.

Development exports point to `src/`, while `publishConfig` maps the same subpaths to `dist/`. `prepack` builds from a clean source tree before packing or publishing.

The one-module-to-one-file ESM output lets compatible bundlers remove unused JavaScript exports. The stylesheet remains listed in `sideEffects` so bundlers do not discard it.

The `files` field limits the package to runtime output and public documentation. `packaging/tarballContents.test.ts` is the authority for what must be present or absent, and [Testing](./testing.md#package-checks) explains the package checks.

## Peer dependencies

ProseMirror packages are peers because objects such as `PluginKey` and `Decoration` depend on module identity. Bundling a second copy would make consumer plugins incompatible with the editor's copy.

A ProseMirror package imported by source code belongs in both `peerDependencies`, for consumers, and `devDependencies`, for this package's tests. The root entry re-exports the common ProseMirror types needed by extension code.

The build leaves all peers external, and the packaging tests ensure published dependency ranges no longer contain workspace catalog syntax.
