/**
 * The notes drawn over the paper, in the room the page layout kept for them.
 *
 * One component draws both kinds, since they differ only in where a page's room comes from, what
 * the area is called, its class, whether the rule above the notes is drawn on that page, and what
 * is drawn in it; the layer over the sheet, the row that takes its room before it is measured, the
 * one editing view and the press that opens a note hold for both.
 *
 * Each area is given the height the layout kept, so one asking for more room than the page had
 * left scrolls inside it.
 */

import type { CSSProperties, ReactElement } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import { FOOTNOTE_BAND } from "../../page/demands/footnoteDemands";
import type { PageOverlay } from "../../page/usePageLayout";
import { noteName, type StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";
import { NOTE_SEPARATOR_HEIGHT, noteSeparatorWidth } from "./noteSeparator";
import { type NoteHeightReport, type RowEditing, StoryRow } from "./StoryRow";

/** The room one page keeps for notes, as the notes drawn in it are placed and named */
export interface NoteArea {
  readonly page: number;
  /** What assistive technology calls the area */
  readonly name: string;
  /** What the page keeps there, by story key, in the order the layout laid them */
  readonly ids: readonly string[];
  readonly top: number;
  readonly left: number;
  readonly width: number;
  readonly height: number;
  /** Whether the notes ask for more room than the page had left, so the area scrolls them */
  readonly scrolls: boolean;
  /** Whether the rule the notes are set under is drawn above this page's */
  readonly rule: boolean;
}

/** The rooms each page keeps for its footnotes, which every page holding one is given */
export function footnoteAreasOf(overlay: PageOverlay): readonly NoteArea[] {
  return overlay.pages.flatMap((page) =>
    page.reserved
      .filter((room) => room.band === FOOTNOTE_BAND)
      .map((room) => ({
        page: page.page,
        name: `Footnotes on page ${page.page}`,
        ids: room.ids,
        top: room.top,
        left: page.left,
        width: page.width,
        height: room.height,
        scrolls: room.clipped,
        // A page's footnotes are set under a rule of their own, wherever they stand
        rule: true,
      }))
  );
}

/**
 * The rooms the endnotes take after the last paragraph of the document.
 *
 * They are one run of notes rather than a page's own, so the rule is drawn once, above the page
 * they begin on.
 */
export function endnoteAreasOf(overlay: PageOverlay): readonly NoteArea[] {
  const opening = overlay.pages.find((page) => page.trailing !== null)?.page;
  return overlay.pages.flatMap((page) => {
    const room = page.trailing;
    return room === null
      ? []
      : [
          {
            page: page.page,
            name: `Endnotes on page ${page.page}`,
            ids: room.ids,
            top: room.top,
            left: page.left,
            width: page.width,
            height: room.height,
            scrolls: room.clipped,
            rule: page.page === opening,
          },
        ];
  });
}

export interface NoteAreasProps {
  /** Where the pages stand, which is the box the areas are placed inside */
  readonly overlay: PageOverlay;
  readonly areas: readonly NoteArea[];
  /** The class each area carries, which is what tells the two kinds apart on the screen */
  readonly areaClassName: string;
  /** The notes these areas draw, by story key, which is what the layout laid out by id */
  readonly rows: ReadonlyMap<string, NoteRow>;
  readonly heights: ReadonlyMap<StoryKey, number>;
  readonly onHeight: NoteHeightReport;
  readonly fontFallbacks: FontFallbacks;
  /** How the document sets its text, which the sheet's own box does not pass down to these */
  readonly textStyle: CSSProperties;
  readonly zoom: number;
  /** The note the caret is in, which is the one row an editing view stands over */
  readonly open?: StoryKey | null;
  /** What mounts that view. Null while no note is open */
  readonly editing?: RowEditing | null;
  /** Whether the open note stands where nothing may be edited */
  readonly readOnly?: boolean;
  /** A value that differs every time the main document changes, which the open row syncs on */
  readonly revision?: unknown;
  /** Called for a press on a note no view stands over yet */
  readonly onOpen?: (key: StoryKey, at: { left: number; top: number }) => void;
  /** Called for a press on the number a note is drawn by, which is the way back to its reference */
  readonly onReturn?: (key: StoryKey) => void;
}

export function NoteAreas({
  overlay,
  areas,
  areaClassName,
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
  onReturn,
}: NoteAreasProps): ReactElement {
  return (
    <div
      className={editorClassNames.noteAreas}
      style={{
        left: `${overlay.left}px`,
        top: `${overlay.top}px`,
        width: `${overlay.width}px`,
        height: `${overlay.sheetHeight}px`,
      }}
    >
      {areas.map((area) => (
        <section
          key={area.page}
          className={areaClassName}
          aria-label={area.name}
          style={{
            ...textStyle,
            top: `${area.top}px`,
            left: `${area.left}px`,
            width: `${area.width}px`,
            height: `${area.height}px`,
            overflowY: area.scrolls ? "auto" : "hidden",
          }}
        >
          {area.rule && (
            <div
              className={editorClassNames.noteSeparator}
              style={{
                height: `${NOTE_SEPARATOR_HEIGHT}px`,
                width: `${noteSeparatorWidth(area.width)}px`,
              }}
            />
          )}
          {area.ids.flatMap((id) => {
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
                // A note nobody has measured yet takes its room unseen, unless it is the one the
                // caret is going into: an unseen row takes no caret
                hidden={!entered && !heights.has(row.key)}
                onHeight={onHeight}
                zoom={zoom}
                editing={entered ? editing : null}
                readOnly={readOnly}
                revision={entered ? revision : undefined}
                onPress={
                  onOpen === undefined ? undefined : (at) => onOpen(row.key, at)
                }
                onReturn={
                  onReturn === undefined ? undefined : () => onReturn(row.key)
                }
              />,
            ];
          })}
        </section>
      ))}
    </div>
  );
}
