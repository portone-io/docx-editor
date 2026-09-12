/**
 * The endnotes after the last paragraph of the document, over the room the page layout kept for
 * them (`page/pageLayout`).
 *
 * Word puts the endnotes at the end of the document rather than at the foot of the page that calls
 * them, so they follow the last block on the page it ends on, under a rule drawn once, and run on
 * down pages of their own. Each one is markup drawn from its story (`./StoryRow`) and each reports
 * the height it is drawn at, so the room the layout kept and the notes drawn in it are the same
 * size.
 *
 * A document asking for its endnotes at the end of each section (`w:endnotePr/w:pos` `sectEnd`)
 * gets them at the end of the document all the same, which is written down as a limit.
 */

import type { CSSProperties, ReactElement } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import type { PageOverlay } from "../../page/usePageLayout";
import { noteName, type StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";
import { NOTE_SEPARATOR_HEIGHT, noteSeparatorWidth } from "./noteSeparator";
import { type NoteHeightReport, type RowEditing, StoryRow } from "./StoryRow";

export interface TrailingNotesProps {
  readonly overlay: PageOverlay;
  /** The notes drawn here, by story key, which is what the layout laid out by id */
  readonly rows: ReadonlyMap<string, NoteRow>;
  readonly heights: ReadonlyMap<StoryKey, number>;
  readonly onHeight: NoteHeightReport;
  readonly fontFallbacks: FontFallbacks;
  /** How the document sets its text, which the sheet's own box does not pass down to these */
  readonly textStyle: CSSProperties;
  readonly zoom: number;
  /** The endnote the caret is in, which is the one row an editing view stands over */
  readonly open?: StoryKey | null;
  /** What mounts that view. Null while no endnote is open */
  readonly editing?: RowEditing | null;
  /** Whether the open endnote stands where nothing may be edited */
  readonly readOnly?: boolean;
  /** A value that differs every time the main document changes, which the open row syncs on */
  readonly revision?: unknown;
  /** Called for a press on an endnote no view stands over yet */
  readonly onOpen?: (key: StoryKey, at: { left: number; top: number }) => void;
}

export function TrailingNotes({
  overlay,
  rows,
  heights,
  onHeight,
  fontFallbacks,
  textStyle,
  zoom,
  open = null,
  editing = null,
  readOnly = false,
  revision,
  onOpen,
}: TrailingNotesProps): ReactElement {
  // The rule is drawn once, above the first of the notes, which is what the layout kept room for
  const opening = overlay.pages.find((page) => page.trailing !== null)?.page;
  return (
    <div
      className={editorClassNames.footnoteAreas}
      style={{
        left: `${overlay.left}px`,
        top: `${overlay.top}px`,
        width: `${overlay.width}px`,
        height: `${overlay.sheetHeight}px`,
      }}
    >
      {overlay.pages.flatMap((page) => {
        const room = page.trailing;
        if (room === null) return [];
        return [
          <section
            key={page.page}
            className={editorClassNames.endnoteArea}
            aria-label={`Endnotes on page ${page.page}`}
            style={{
              ...textStyle,
              top: `${room.top}px`,
              left: `${page.left}px`,
              width: `${page.width}px`,
            }}
          >
            {page.page === opening && (
              <div
                className={editorClassNames.noteSeparator}
                style={{
                  height: `${NOTE_SEPARATOR_HEIGHT}px`,
                  width: `${noteSeparatorWidth(page.width)}px`,
                }}
              />
            )}
            {room.ids.flatMap((id) => {
              const row = rows.get(id);
              if (row === undefined) return [];
              const entered = row.key === open;
              return [
                <StoryRow
                  key={row.key}
                  noteKey={row.key}
                  story={row.story}
                  label={row.label}
                  name={noteName(row.kind, row.label)}
                  fontFallbacks={fontFallbacks}
                  // A note nobody has measured yet takes its room unseen, unless it is the one
                  // the caret is going into: an unseen row takes no caret
                  hidden={!entered && !heights.has(row.key)}
                  onHeight={onHeight}
                  zoom={zoom}
                  editing={entered ? editing : null}
                  readOnly={readOnly}
                  revision={entered ? revision : undefined}
                  onPress={
                    onOpen === undefined
                      ? undefined
                      : (at) => onOpen(row.key, at)
                  }
                />,
              ];
            })}
          </section>,
        ];
      })}
    </div>
  );
}
