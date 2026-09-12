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
import type { PagePixels, TrailingRows } from "../../page/pageLayout";
import type { PageOverlay } from "../../page/usePageLayout";
import type { StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../../styles/fontStack";
import { documentDefaultsVariables } from "../../styles/inlineStyle";
import { endnoteAreasOf, footnoteAreasOf, NoteAreas } from "./NoteAreas";
import { NoteList } from "./NoteList";
import { useNoteHeights } from "./noteHeights";
import { NOTE_SEPARATOR_HEIGHT } from "./noteSeparator";
import type { NoteHeightReport, RowEditing } from "./StoryRow";

const NO_ROWS: ReadonlyMap<StoryKey, NoteRow> = new Map();
const NO_NOTE_ROWS: readonly NoteRow[] = [];

/** What a document's notes ask of the page layout, and what is needed to draw them */
export interface NoteBands {
  /**
   * The bands to lay the pages out with, by name, or undefined where the document asks for no
   * room at all. A new value lays the pages out again, so it changes only when a height does
   */
  readonly bands: ReadonlyMap<string, DemandBand> | undefined;
  /**
   * The endnotes as rows laid after the last block of the document, or undefined where it refers
   * to none. A new value lays the pages out again, so it changes only when a height does
   */
  readonly trailing: TrailingRows | undefined;
  /**
   * Every note the document refers to, by story key, in first-reference order.
   *
   * Which one a caret stands in is the mounting component's to hold, so this is the one half of
   * the notes it reads; everything else about how they are drawn stays in `drawn`.
   */
  readonly rows: ReadonlyMap<StoryKey, NoteRow>;
  /** Held for `NotesAroundPage`; nothing else reads it */
  readonly drawn: DrawnNotes;
}

/** What `NotesAroundPage` draws with, kept opaque so a caller cannot take the two halves apart */
interface DrawnNotes {
  /** The footnotes, which are the notes drawn at the foot of the page their reference stands on */
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
  const footnotes = notes?.footnotes ?? NO_ROWS;
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

  // The endnotes stand after the last paragraph rather than beside the place that calls them, so
  // they are rows the layout lays last rather than a band a page keeps at its foot
  const endnotes = notes?.endnotes ?? NO_NOTE_ROWS;
  const trailing = useMemo<TrailingRows | undefined>(
    () =>
      endnotes.length === 0
        ? undefined
        : {
            overhead: NOTE_SEPARATOR_HEIGHT,
            rows: endnotes.map((row) => ({
              id: row.key,
              height: heights.get(row.key) ?? 0,
            })),
          },
    [endnotes, heights]
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
    trailing,
    rows: notes?.rows ?? NO_ROWS,
    drawn: {
      footnotes,
      endnotes,
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
  /**
   * Where the pages stand, or null while none are measured, which is what decides where the notes
   * go: a page keeps room for them only once it has been laid out
   */
  readonly overlay: PageOverlay | null;
  /** The paper the first section is drawn on, which the list under the sheet is as wide as */
  readonly page: PagePixels;
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
}

/**
 * The notes around the paper: each footnote over the room its page keeps, and the endnotes over
 * the room kept after the last paragraph.
 *
 * Until a page has been measured there is no room to stand in - the guides may be off, or the
 * first frame may not have been laid out yet - and a note drawn nowhere is a note a reader cannot
 * read, so both kinds are listed under the sheet until there is. Only a note standing in its own
 * room is edited in place.
 */
export function NotesAroundPage({
  notes,
  overlay,
  page,
  zoom,
  open,
  editing,
  readOnly,
  revision,
  onOpen,
}: NotesAroundPageProps): ReactElement | null {
  const { footnotes, endnotes, anyNotes, heights, onHeight } = notes.drawn;
  const drawing = {
    heights,
    onHeight,
    fontFallbacks: notes.drawn.fontFallbacks,
    textStyle: notes.drawn.textStyle,
    zoom,
    open,
    editing,
    readOnly,
    revision,
    onOpen,
  };
  return (
    <>
      {overlay !== null && footnotes.size > 0 && (
        <NoteAreas
          overlay={overlay}
          areas={footnoteAreasOf(overlay)}
          areaClassName={editorClassNames.footnoteArea}
          rows={footnotes}
          {...drawing}
        />
      )}
      {overlay !== null && endnotes.length > 0 && (
        <NoteAreas
          overlay={overlay}
          areas={endnoteAreasOf(overlay)}
          areaClassName={editorClassNames.endnoteArea}
          rows={notes.rows}
          {...drawing}
        />
      )}
      {anyNotes && overlay === null && (
        <NoteList
          footnotes={[...footnotes.values()]}
          endnotes={endnotes}
          page={page}
          fontFallbacks={notes.drawn.fontFallbacks}
          textStyle={notes.drawn.textStyle}
        />
      )}
    </>
  );
}
