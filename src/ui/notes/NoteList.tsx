/**
 * The notes listed after the last page: every endnote, and the footnotes ahead of them while no
 * pages are drawn for footnotes to stand at the foot of.
 *
 * Each note is drawn from its story (`./StoryRow`) across the width the body is set at, so a note
 * wraps here where it would on the page.
 */

import type { CSSProperties, ReactElement } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import type { PagePixels } from "../../page/pageLayout";
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
          fontFallbacks={fontFallbacks}
        />
      ))}
    </section>
  );
}
