# Table cell layout

## Vertical alignment

`w:vAlign` places a cell's content at `top`, `center`, or `bottom`. An omitted value displays as top. Although the shared value type also contains `both`, ECMA-376 requires consumers to ignore it on a table cell and return to top. The editor writes a direct `w:vAlign` only on selected cells and retains unrelated `w:tcPr` children and attributes.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.4.83, 17.18.101.

## Cell margins

`w:tblCellMar` in direct table properties supplies table-level cell margin defaults. A side declared by `w:tcMar` on a cell overrides that side only; all other sides continue to inherit. The editor resolves table-style and direct-table defaults for display and writes direct `dxa` margins only for the sides the user changes. Row-level `w:tblPrEx` margin exceptions remain preserved but are not resolved for display.

Values are exposed in points and serialized in twentieths of a point. Strict leading and trailing margins are read as left and right for the left-to-right display model, with Transitional `left` and `right` spellings as fallbacks. Right-to-left table layout remains preservation-only. A command may target one cell or a rectangular cell selection; mixed values are reported per side. Content-locked cells refuse layout commands together with their UI controls.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §§17.4.42, 17.4.68.

## Collapsed cell borders

When cell spacing is zero, a direct cell border wins a conflict with a table border. For browser
display, the editor therefore suppresses the inherited opposing side and leaves the direct border
as the only collapsed-border candidate. This also avoids spreading one segment's color across the
whole side of an adjacent merged cell. Export still writes the direct border only on the cell the
user formatted. Conflicts between two direct cell borders remain subject to the standard's border
weight, style, color, and reading-order rules.

Observed 2026-08-23 against ECMA-376 5th edition, Part 1, §17.4.66.

## Row height

`w:trHeight w:val` is measured in twentieths of a point. A pointer resize writes an explicit
`atLeast` rule for a row that had no exact constraint, allowing content to make the row taller, and
retains `exact` when the document already required it. The edit changes only the target row and
preserves its other `w:trPr` properties.

Observed 2026-08-23 against ECMA-376 5th edition, Part 1, §17.4.80.
