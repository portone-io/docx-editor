# Preservation tiers

## What a tier is

WordprocessingML admits far more elements at each level than an editor models. What a reader does with the rest decides how much of the file survives an edit, so the answer is written down once, per level, rather than being implied by which cases a `switch` happens to cover.

Every element the readers may meet carries one of six tiers.

| Tier | What it means |
| --- | --- |
| `model` | A reader turns it into an editable node. |
| `marker` | One end of a range: invisible, and its pair and its order must survive every edit. |
| `ignorable` | A trace the producer regenerates: invisible, and an edit may drop it. |
| `runContent` | A `CT_R` child kept verbatim inside its run, wearing the run's formatting. |
| `inline` | A paragraph child kept verbatim beside the runs. |
| `block` | A body or cell child kept verbatim as a block placeholder. |

Beside the tier, a rule says how the element is drawn (`hidden`, `text`, `break`, `chip`) and whether the deletion guard answers for it.

## Demotion is a structure failure, not a vocabulary

The tiers are grouped by the content model each level has in `wml.xsd`, one sub-table per level: `body`, `tbl`, `tr`, `tc`, `p`, `wrapper`, `r`. An element a sub-table does not name falls to the narrowest preservation that level has a node for. A run child stays inside its run, a paragraph child beside the runs, a block stays a block, so an element nobody modelled costs that element and nothing around it, whatever depth it was met at.

Two levels are the exception. `tbl` holds rows and `tr` holds cells, and neither has a node a stranger could be kept in, so an element they do not name stands the whole table down as one placeholder and the fidelity report says so with `table-demoted`.

What those two levels do name as `marker` or `ignorable` is invisible, so it needs no node of its own: it is carried on the child before it - a row's on the cell it followed, a table's on the row - and written back in the same spot. The one case that still stands the table down is a marker following a cell that only continues a vertical merge, since that cell is created fresh on export and would have nothing to carry it.

The level rather than the element name is what the sub-tables are keyed by, because the same name means different things in different places. `w:sdt` is a content control a paragraph reader unwraps, a wrapper around one cell under `w:tr`, a block placeholder under `w:body`, and a row wrapper nothing reads under `w:tbl`. Keys are Clark names (`{namespace}localName`) so that `m:oMath`, which is not WordprocessingML at all, sits in the same map.

## Why `w:lastRenderedPageBreak` is ignorable

Part 1 §17.3.3.13 defines `w:lastRenderedPageBreak` as the position at which the producer's own layout last broke the page. It is a cache of a rendering, not content the document asserts, and a producer rewrites it whenever it lays the document out again. It is therefore kept where it stood and drawn as nothing, and an edit that drops it is refused by nothing: the file says the same thing without it.

This matters more than it sounds. Word writes one on every page of a document it has laid out, so a reader that stood a paragraph down for an element it did not model stood down one paragraph per page of every real file.

## Why the pieces of a field are guarded

Part 1 §17.16.18 defines a complex field as a `w:fldChar` of type `begin`, the instruction that follows it as `w:instrText`, an optional `w:fldChar` of type `separate`, the result last drawn for it, and a `w:fldChar` of type `end`. The field is those pieces read in that order; one of them taken away leaves the rest saying something the file never said, and nothing this editor writes can put such a piece back.

They are therefore `guarded`, alongside the range markers whose two ends a document falls apart without: `w:bookmarkStart`/`w:bookmarkEnd` (§17.13.6), `w:permStart`/`w:permEnd` (§17.13.7), and the move ranges of §17.13.5. For preserved nodes the guard reads that attribute. Markers carried in a table, row or cell attribute have no node of their own: the guard reads their XML using the same range-marker vocabulary as the policy (`src/ooxml/rangeMarkers.ts`), ignoring producer traces such as `proofErr`. It compares those markers and guarded nodes together in document order.

A container that carries its content whole is not guarded. A `w:ins`, a `w:del`, a `w:fldSimple` or a `w:smartTag` says everything it says inside itself, so removing it removes a self-contained piece of the document and leaves nothing dangling. Those are drawn as chips a reader can select and delete.

## What each level's table was written from

| Level | Content model |
| --- | --- |
| `r` | `CT_R`: `EG_RPr` and `EG_RunInnerContent` |
| `p` | `CT_P`: `w:pPr` and `EG_PContent` |
| `wrapper` | `CT_SdtContentRun` and `CT_Hyperlink`, both `EG_PContent` |
| `body` | `CT_Body`: `EG_BlockLevelElts` and the closing `w:sectPr` |
| `tc` | `CT_Tc`: `w:tcPr` and `EG_BlockLevelElts` |
| `tbl` | `CT_Tbl`: `EG_RangeMarkupElements`, `w:tblPr`, `w:tblGrid`, `EG_ContentRowContent` |
| `tr` | `CT_Row`: `w:tblPrEx`, `w:trPr`, `EG_ContentCellContent` |

`src/docx/importPolicy.test.ts` expands each of those types out of the committed schema, groups and cross-schema references followed, and compares the result with the sub-table in both directions. An element the schema admits at a level and the table does not name fails there, except for the ones `tbl` and `tr` demote, which the test names one by one.

`EG_CellMarkupElements` (`w:cellIns`, `w:cellDel`, `w:cellMerge`) is not a row child. `wml.xsd` references it only from `CT_TcPr`, so those three travel inside a cell's properties rather than at the level a reader walks.
