/** Grid for choosing the dimensions of a table to insert. */

import { useRef, useState } from "react";
import { insertTable } from "../editor/insertTable";
import { editorClassNames } from "../styles/classNames";
import type { RunCommand } from "./runCommand";
import { useGridKeyboard } from "./useGridKeyboard";

const MAX_SIDE = 6;

const SIDES = Array.from({ length: MAX_SIDE }, (_, index) => index + 1);

interface Size {
  rows: number;
  cols: number;
}

const NO_SIZE: Size = { rows: 0, cols: 0 };

export interface TableSizePickerProps {
  run: RunCommand;
  close: () => void;
  takeFocus: boolean;
}

export function TableSizePicker({
  run,
  close,
  takeFocus,
}: TableSizePickerProps) {
  const [size, setSize] = useState<Size>(NO_SIZE);
  const grid = useRef<HTMLDivElement | null>(null);
  const keys = useGridKeyboard({ grid, onClose: close, takeFocus });

  const choose = (rows: number, cols: number) => {
    run(insertTable({ rows, columns: cols }));
    close();
  };

  return (
    <>
      <div
        ref={grid}
        className={editorClassNames.sizeGrid}
        role="grid"
        aria-label="Table size"
        {...keys}
      >
        {SIDES.map((rows) => (
          <div key={rows} className={editorClassNames.sizeGridRow} role="row">
            {SIDES.map((cols) => (
              <button
                key={cols}
                type="button"
                role="gridcell"
                className={editorClassNames.sizeGridCell}
                aria-label={`${rows} by ${cols} table`}
                data-on={
                  rows <= size.rows && cols <= size.cols ? "" : undefined
                }
                onMouseDown={(event) => event.preventDefault()}
                onMouseEnter={() => setSize({ rows, cols })}
                onFocus={() => setSize({ rows, cols })}
                onClick={() => choose(rows, cols)}
              />
            ))}
          </div>
        ))}
      </div>
      <p className={editorClassNames.sizeGridLabel} aria-live="polite">
        {size.rows > 0 ? `${size.rows} x ${size.cols}` : "Pick a size"}
      </p>
    </>
  );
}
