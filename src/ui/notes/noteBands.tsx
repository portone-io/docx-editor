/**
 * The footnote band: what a page keeps room for at its foot, and what is drawn in that room and
 * after the last page.
 *
 * A band is two halves that have to agree. The layout is told a name, a place at the foot, the
 * height a page adds once for the band and the height of each thing kept in it (`page/demands`),
 * and something has to draw exactly that much. Both halves live here so a page's room and the
 * notes drawn in it cannot drift apart, and so the component that mounts the editor asks only for
 * the bands a document keeps and renders what it is handed.
 *
 * A document referring to no footnote keeps no band at all, so it is laid out through the same
 * pass it took before a page kept room for anything.
 */

import type { EditorState } from "prosemirror-state";
import type { CSSProperties, ReactElement } from "react";
import { useMemo } from "react";
import {
  type NoteRow,
  noteProjection,
} from "../../editor/commands/noteQueries";
import { documentOf } from "../../editor/editorDocument";
import type { DemandBand } from "../../page/demands";
import {
  FOOTNOTE_BAND,
  FOOTNOTE_BAND_ORDER,
} from "../../page/demands/footnoteDemands";
import type { PagePixels } from "../../page/pageLayout";
import type { PageOverlay } from "../../page/usePageLayout";
import type { StoryKey } from "../../schema/stories";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../../styles/fontStack";
import { documentDefaultsVariables } from "../../styles/inlineStyle";
import { FootnoteAreas } from "./FootnoteAreas";
import { NoteList } from "./NoteList";
import { useNoteHeights } from "./noteHeights";
import { NOTE_SEPARATOR_HEIGHT } from "./noteSeparator";
import type { NoteHeightReport } from "./StoryRow";

const NO_FOOTNOTES: ReadonlyMap<StoryKey, NoteRow> = new Map();
const NO_NOTE_ROWS: readonly NoteRow[] = [];

/** What a document's notes ask of the page layout, and what is needed to draw them */
export interface NoteBands {
  /**
   * The bands to lay the pages out with, by name, or undefined where the document asks for no
   * room at all. A new value lays the pages out again, so it changes only when a height does
   */
  readonly bands: ReadonlyMap<string, DemandBand> | undefined;
  /** Held for `NotesAroundPage`; nothing else reads it */
  readonly drawn: DrawnNotes;
}

/** What `NotesAroundPage` draws with, kept opaque so a caller cannot take the two halves apart */
interface DrawnNotes {
  readonly footnotes: ReadonlyMap<StoryKey, NoteRow>;
  readonly endnotes: readonly NoteRow[];
  readonly anyNotes: boolean;
  readonly heights: ReadonlyMap<StoryKey, number>;
  readonly onHeight: NoteHeightReport;
  readonly fontFallbacks: FontFallbacks;
  readonly textStyle: CSSProperties;
}

export interface NoteBandsOptions {
  /** The live document's state, or null before one is opened */
  readonly state: EditorState | null;
  /** A value that differs for every document opened, which the measured heights are held against */
  readonly of: unknown;
  readonly fontFallbacks: FontFallbacks | undefined;
}

export function useNoteBands({
  state,
  of,
  fontFallbacks,
}: NoteBandsOptions): NoteBands {
  const notes = state === null ? null : noteProjection.read(state);
  const footnotes = notes?.footnotes ?? NO_FOOTNOTES;
  const hasFootnotes = footnotes.size > 0;
  const [heights, onHeight] = useNoteHeights(of);
  const bands = useMemo(
    () =>
      hasFootnotes
        ? new Map<string, DemandBand>([
            [
              FOOTNOTE_BAND,
              {
                order: FOOTNOTE_BAND_ORDER,
                overhead: NOTE_SEPARATOR_HEIGHT,
                heights,
              },
            ],
          ])
        : undefined,
    [hasFootnotes, heights]
  );

  const snapshot = state === null ? null : documentOf(state);
  const noteFontFallbacks = fontFallbacks ?? DEFAULT_FONT_FALLBACKS;
  // The sheet sets its text from variables on its own box, which the notes drawn beside it are
  // not inside of
  const textStyle = useMemo<CSSProperties>(
    () =>
      snapshot === null
        ? {}
        : {
            ...documentDefaultsVariables(snapshot.defaults, noteFontFallbacks),
            tabSize: `${snapshot.defaultTabStopPt}pt`,
          },
    [snapshot, noteFontFallbacks]
  );

  return {
    bands,
    drawn: {
      footnotes,
      endnotes: notes?.endnotes ?? NO_NOTE_ROWS,
      anyNotes: notes !== null,
      heights,
      onHeight,
      fontFallbacks: noteFontFallbacks,
      textStyle,
    },
  };
}

export interface NotesAroundPageProps {
  readonly notes: NoteBands;
  /** Where the pages stand, or null while none are drawn */
  readonly overlay: PageOverlay | null;
  /** The paper the first section is drawn on, which the list after the last page is as wide as */
  readonly page: PagePixels;
  /** Whether the pages are drawn at all, which is what decides where the footnotes go */
  readonly pageGuides: boolean;
  readonly zoom: number;
}

/**
 * The notes around the paper: each footnote over the room its page keeps, and the notes listed
 * after the last page. With no pages drawn there is no room to stand in, so the footnotes join
 * that list ahead of the endnotes instead.
 */
export function NotesAroundPage({
  notes,
  overlay,
  page,
  pageGuides,
  zoom,
}: NotesAroundPageProps): ReactElement | null {
  const { footnotes, endnotes, anyNotes, heights, onHeight } = notes.drawn;
  return (
    <>
      {overlay !== null && footnotes.size > 0 && (
        <FootnoteAreas
          overlay={overlay}
          footnotes={footnotes}
          heights={heights}
          onHeight={onHeight}
          fontFallbacks={notes.drawn.fontFallbacks}
          textStyle={notes.drawn.textStyle}
          zoom={zoom}
        />
      )}
      {anyNotes && (
        <NoteList
          footnotes={pageGuides ? NO_NOTE_ROWS : [...footnotes.values()]}
          endnotes={endnotes}
          page={page}
          fontFallbacks={notes.drawn.fontFallbacks}
          textStyle={notes.drawn.textStyle}
        />
      )}
    </>
  );
}
