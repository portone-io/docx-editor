/** Optional picker values; omitted fields use the package's exported defaults. */

import type { ColorRow } from "../styles/colors";
import type { LineSpacingOption } from "../styles/lineSpacings";
import type { CellBorderOption } from "../table/cellBorders";

export interface DocxEditorPresets {
  fonts?: readonly string[];
  colors?: readonly ColorRow[];
  fontSizes?: readonly number[];
  lineSpacings?: readonly LineSpacingOption[];
  cellBorders?: readonly CellBorderOption[];
}
