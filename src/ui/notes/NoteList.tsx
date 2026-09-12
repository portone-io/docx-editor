/**
 * The notes listed under the sheet while no page has kept room for them: the footnotes ahead of
 * the endnotes, the order they are called in.
 *
 * A note stands where its own page keeps room for it (`./NoteAreas`), which needs a page that has
 * been laid out: the guides may be off, or the first frame may not have been measured yet. Until
 * then they are listed here, where none of them is edited in place.
 *
 * Each note is drawn from its story (`./StoryRow`) across the width the body is set at, so a note
 * wraps here where it would on the page.
 */

import type { CSSProperties, ReactElement } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import type { PagePixels } from "../../page/pageLayout";
import { noteName } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";
import { NOTE_SEPARATOR_HEIGHT, noteSeparatorWidth } from "./noteSeparator";
import { StoryRow } from "./StoryRow";

export interface NoteListProps {
  readonly footnotes: readonly NoteRow[];
  readonly endnotes: readonly NoteRow[];
  /** The paper the first section is drawn on, which the list is as wide as */
  readonly page: PagePixels;
  readonly fontFallbacks: FontFallbacks;
  /** How the document sets its text, which the sheet's own box does not pass down to this */
  readonly textStyle: CSSProperties;
}

function listName(footnotes: boolean, endnotes: boolean): string {
  if (footnotes && endnotes) return "Footnotes and endnotes";
  return footnotes ? "Footnotes" : "Endnotes";
}

export function NoteList({
  footnotes,
  endnotes,
  page,
  fontFallbacks,
  textStyle,
}: NoteListProps): ReactElement | null {
  if (footnotes.length === 0 && endnotes.length === 0) return null;
  return (
    <section
      className={editorClassNames.noteList}
      aria-label={listName(footnotes.length > 0, endnotes.length > 0)}
      style={{
        ...textStyle,
        width: `${page.pageWidth}px`,
        paddingLeft: `${page.marginLeft}px`,
        paddingRight: `${page.marginRight}px`,
      }}
    >
      <div
        className={editorClassNames.noteSeparator}
        style={{
          height: `${NOTE_SEPARATOR_HEIGHT}px`,
          width: `${noteSeparatorWidth(page.bodyWidth)}px`,
        }}
      />
      {[...footnotes, ...endnotes].map((row) => (
        <StoryRow
          key={row.key}
          noteKey={row.key}
          story={row.story}
          label={row.label}
          name={noteName(row.kind, row.label)}
          fontFallbacks={fontFallbacks}
        />
      ))}
    </section>
  );
}
