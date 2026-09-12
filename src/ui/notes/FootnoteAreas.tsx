/**
 * The footnotes at the foot of each page, over the room the page layout keeps for them.
 *
 * Every footnote is markup drawn from its story (`./StoryRow`), and each page's footnotes stand
 * under a short rule in the order the page keeps them. The rule's height is what the layout is
 * told a page adds once when it holds any footnote, and each row reports the height it is drawn
 * at, so the room kept and the notes drawn in it are the same size. A row not yet measured takes
 * its room unseen, and a page whose footnotes are taller than it can hold scrolls them.
 */

import type { CSSProperties, ReactElement } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import { FOOTNOTE_BAND } from "../../page/demands/footnoteDemands";
import type { PageOverlay } from "../../page/usePageLayout";
import type { StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";
import { NOTE_SEPARATOR_HEIGHT, noteSeparatorWidth } from "./noteSeparator";
import { type NoteHeightReport, StoryRow } from "./StoryRow";

export interface FootnoteAreasProps {
  readonly overlay: PageOverlay;
  readonly footnotes: ReadonlyMap<string, NoteRow>;
  readonly heights: ReadonlyMap<StoryKey, number>;
  readonly onHeight: NoteHeightReport;
  readonly fontFallbacks: FontFallbacks;
  /** How the document sets its text, which the sheet's own box does not pass down to these */
  readonly textStyle: CSSProperties;
  readonly zoom: number;
}

export function FootnoteAreas({
  overlay,
  footnotes,
  heights,
  onHeight,
  fontFallbacks,
  textStyle,
  zoom,
}: FootnoteAreasProps): ReactElement {
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
      {overlay.pages.flatMap((page) =>
        page.reserved
          .filter((room) => room.band === FOOTNOTE_BAND)
          .map((room) => (
            <section
              key={page.page}
              className={editorClassNames.footnoteArea}
              aria-label={`Footnotes on page ${page.page}`}
              style={{
                ...textStyle,
                top: `${room.top}px`,
                left: `${page.left}px`,
                width: `${page.width}px`,
                height: `${room.height}px`,
                overflowY: room.clipped ? "auto" : "hidden",
              }}
            >
              <div
                className={editorClassNames.noteSeparator}
                style={{
                  height: `${NOTE_SEPARATOR_HEIGHT}px`,
                  width: `${noteSeparatorWidth(page.width)}px`,
                }}
              />
              {room.ids.flatMap((id) => {
                const row = footnotes.get(id);
                return row === undefined
                  ? []
                  : [
                      <StoryRow
                        key={row.key}
                        noteKey={row.key}
                        story={row.story}
                        label={row.label}
                        fontFallbacks={fontFallbacks}
                        hidden={!heights.has(row.key)}
                        onHeight={onHeight}
                        zoom={zoom}
                      />,
                    ];
              })}
            </section>
          ))
      )}
    </div>
  );
}
