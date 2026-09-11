import { useCallback, useState } from "react";
import type { StoryKey } from "../../schema/stories";
import type { NoteHeightReport } from "./StoryRow";

const NO_HEIGHTS: ReadonlyMap<StoryKey, number> = new Map();

/** Rounded as the layout rounds, so a height wobbling below the pixel does not lay the pages out again */
function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/**
 * The height each note stands at, as its row last reported it, for the document `of` names.
 *
 * A height is a fact about what the browser drew, so it can only arrive after a row is drawn. The
 * map handed back is the same object until a height really changes, which is what lets the page
 * layout that reads it run again only then, and a document swapped in starts from none.
 */
export function useNoteHeights(
  of: unknown
): readonly [ReadonlyMap<StoryKey, number>, NoteHeightReport] {
  const [held, setHeld] = useState<{
    readonly of: unknown;
    readonly heights: ReadonlyMap<StoryKey, number>;
  }>({ of, heights: NO_HEIGHTS });

  const report = useCallback<NoteHeightReport>(
    (key, height) => {
      const rounded = round(height);
      setHeld((previous) => {
        const heights = previous.of === of ? previous.heights : NO_HEIGHTS;
        if (heights.get(key) === rounded && previous.of === of) return previous;
        return { of, heights: new Map(heights).set(key, rounded) };
      });
    },
    [of]
  );

  return [held.of === of ? held.heights : NO_HEIGHTS, report];
}
