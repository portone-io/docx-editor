/**
 * The rule notes are set under, wherever they are drawn: in the room a page keeps for them
 * (`./NoteAreas`) and at the head of the list under the sheet (`./NoteList`).
 *
 * Its height is also what the page layout is told a page adds once when it holds any footnote
 * (`DemandBand.overhead`), and what the endnotes are given once where they begin
 * (`TrailingRows.overhead`), so the room kept and the notes drawn in it are the same size.
 */

/** The height a page's notes are set under, the rule drawn across the middle of it */
export const NOTE_SEPARATOR_HEIGHT = 16;

/** Word's rule runs a third of the body and no further than two inches */
export function noteSeparatorWidth(bodyWidth: number): number {
  return Math.min(bodyWidth / 3, 192);
}
