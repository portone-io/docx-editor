/**
 * One note drawn from its story, as markup no editor view stands behind (`editor/stories`).
 *
 * A row is mounted again whenever its note moves to another page, and a story no edit touched is
 * the same node afterwards (`schema/stories`), so the markup is kept against the story node rather
 * than against the row: only a note whose story or label changed is drawn again.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  memo,
  type ReactElement,
  useEffect,
  useLayoutEffect,
  useRef,
} from "react";
import { noteNodeSpecs } from "../../editor/notes/noteSurface";
import {
  type StoryMarkupOptions,
  storyMarkup,
} from "../../editor/stories/storyMarkup";
import type { StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";

export type DrawStory = (
  story: PMNode,
  options: StoryMarkupOptions
) => DocumentFragment;

interface Drawn {
  readonly label: string;
  readonly fontFallbacks: FontFallbacks;
  readonly draw: DrawStory;
  readonly markup: DocumentFragment;
}

const drawn = new WeakMap<PMNode, Drawn>();

function markupOf(
  story: PMNode,
  label: string,
  fontFallbacks: FontFallbacks,
  draw: DrawStory
): DocumentFragment {
  const held = drawn.get(story);
  if (
    held?.label === label &&
    held.fontFallbacks === fontFallbacks &&
    held.draw === draw
  ) {
    return held.markup;
  }
  const markup = draw(story, {
    fontFallbacks,
    nodeSpecs: noteNodeSpecs(() => label),
  });
  drawn.set(story, { label, fontFallbacks, draw, markup });
  return markup;
}

/** Hands on the height a row stands at on the paper, the zoom taken back out */
export type NoteHeightReport = (key: StoryKey, height: number) => void;

export interface StoryRowProps {
  readonly noteKey: StoryKey;
  readonly story: PMNode;
  readonly label: string;
  readonly fontFallbacks: FontFallbacks;
  /** A hidden row still takes its room, which is what lets it be measured before it is placed */
  readonly hidden?: boolean;
  /** Called whenever the row's height changes. Left out, the row is not measured */
  readonly onHeight?: NoteHeightReport;
  /** The visual scale the row is drawn at, which a height read off the screen is divided by */
  readonly zoom?: number;
  /** What draws the story; `storyMarkup` unless a test counts the drawings */
  readonly draw?: DrawStory;
}

export const StoryRow = memo(function StoryRow({
  noteKey,
  story,
  label,
  fontFallbacks,
  hidden = false,
  onHeight,
  zoom = 1,
  draw = storyMarkup,
}: StoryRowProps): ReactElement {
  const box = useRef<HTMLDivElement | null>(null);
  const scale = useRef(zoom);

  useEffect(() => {
    scale.current = zoom;
  });

  // The markup is the story's and not React's, so it is put in after the row is drawn, and
  // before the browser paints
  useLayoutEffect(() => {
    box.current?.replaceChildren(
      markupOf(story, label, fontFallbacks, draw).cloneNode(true)
    );
  }, [story, label, fontFallbacks, draw]);

  useEffect(() => {
    const row = box.current;
    // Where there is no layout (in tests) nothing is ever resized either
    if (!row || !onHeight || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(() => {
      onHeight(noteKey, row.getBoundingClientRect().height / scale.current);
    });
    observer.observe(row);
    return () => observer.disconnect();
  }, [noteKey, onHeight]);

  return (
    <div
      ref={box}
      className={editorClassNames.noteRow}
      style={hidden ? { visibility: "hidden" } : undefined}
    />
  );
});
