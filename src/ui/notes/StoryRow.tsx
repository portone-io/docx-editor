/**
 * One note drawn from its story: as markup no editor view stands behind (`editor/stories`), and as
 * the one editing view while the caret is in it.
 *
 * A row is mounted again whenever its note moves to another page, and a story no edit touched is
 * the same node afterwards (`schema/stories`), so the markup is kept against the story node rather
 * than against the row: only a note whose story or label changed is drawn again.
 *
 * A row carrying the editing view is mounted again by such a move too, so where the caret stood is
 * handed over in `editing.caret` and the view opens there. The page is not laid out again while
 * that view is composing (`DocxEditor`), so a move never reaches an open composition.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import {
  memo,
  type ReactElement,
  useEffect,
  useLayoutEffect,
  useRef,
} from "react";
import type { EditorDocument } from "../../editor/editorDocument";
import { noteNodeSpecs } from "../../editor/notes/noteSurface";
import {
  type StoryMarkupOptions,
  storyMarkup,
} from "../../editor/stories/storyMarkup";
import {
  caretOf,
  createStoryView,
  type StoryCaret,
  type StoryExtensions,
  type StoryHost,
  type StoryView,
} from "../../editor/stories/storyView";
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

/**
 * What mounts the editing view over this row, which only the row the caret is in is given.
 *
 * It is held apart from the row's own props because a row is drawn again on every change to the
 * document while the view over it has to stand: every field here is the same object for as long as
 * one note is open, and what changes per edit reaches the view through `revision`.
 */
export interface RowEditing {
  readonly host: StoryHost;
  /** The values the story is edited against (`editor/editorDocument`) */
  readonly document: EditorDocument;
  readonly extensions: StoryExtensions;
  /** Where the caret opens, and where it stood when the row was last taken down */
  readonly caret: {
    take(): StoryCaret | null;
    keep(caret: StoryCaret): void;
  };
  readonly onStateChange: (state: EditorState) => void;
}

export interface StoryRowProps {
  readonly noteKey: StoryKey;
  readonly story: PMNode;
  readonly label: string;
  readonly fontFallbacks: FontFallbacks;
  /** What assistive technology calls the row */
  readonly name?: string;
  /** A hidden row still takes its room, which is what lets it be measured before it is placed */
  readonly hidden?: boolean;
  /** Called whenever the row's height changes. Left out, the row is not measured */
  readonly onHeight?: NoteHeightReport;
  /** The visual scale the row is drawn at, which a height read off the screen is divided by */
  readonly zoom?: number;
  /** What draws the story; `storyMarkup` unless a test counts the drawings */
  readonly draw?: DrawStory;
  /** Mounted over this row while the caret is in it, and null while the row is markup */
  readonly editing?: RowEditing | null;
  /** A value that differs every time the main document changes, which is when the view is synced */
  readonly revision?: unknown;
  /** Whether the note stands where nothing may be edited, which a reader is told rather than shown */
  readonly readOnly?: boolean;
  /** Called for a press on a row no view stands over yet, with where on the screen it landed */
  readonly onPress?: (at: { left: number; top: number }) => void;
  /**
   * Called for a press on the number the note is drawn by, which is the way back to the text that
   * calls it rather than a place to put the caret
   */
  readonly onReturn?: () => void;
}

/** Whether a press landed on the number the note is drawn by rather than on what a reader wrote */
function pressedTheNumber(target: EventTarget | null): boolean {
  return (
    target instanceof Element &&
    target.closest(`.${editorClassNames.noteMark}`) !== null
  );
}

export const StoryRow = memo(function StoryRow({
  noteKey,
  story,
  label,
  fontFallbacks,
  name,
  hidden = false,
  onHeight,
  zoom = 1,
  draw = storyMarkup,
  editing = null,
  revision,
  readOnly = false,
  onPress,
  onReturn,
}: StoryRowProps): ReactElement {
  const box = useRef<HTMLDivElement | null>(null);
  const scale = useRef(zoom);
  const live = useRef<StoryView | null>(null);

  useEffect(() => {
    scale.current = zoom;
  });

  // The markup is the story's and not React's, so it is put in after the row is drawn, and
  // before the browser paints. A row the editing view stands over draws none of it
  useLayoutEffect(() => {
    if (editing) return;
    box.current?.replaceChildren(
      markupOf(story, label, fontFallbacks, draw).cloneNode(true)
    );
  }, [editing, story, label, fontFallbacks, draw]);

  // biome-ignore lint/correctness/useExhaustiveDependencies: the story and the label are what the view reads out of the main document for itself, so a change in either must not take the view down
  useLayoutEffect(() => {
    const row = box.current;
    if (!row || !editing) return;
    const mount = document.createElement("div");
    row.replaceChildren(mount);
    const opened = createStoryView({
      mount,
      host: editing.host,
      key: noteKey,
      document: editing.document,
      fontFallbacks,
      extensions: editing.extensions,
      caret: editing.caret.take(),
      onStateChange: editing.onStateChange,
    });
    live.current = opened;
    opened.view.focus();
    return () => {
      editing.caret.keep(caretOf(opened.view.state));
      live.current = null;
      opened.destroy();
    };
  }, [editing, noteKey, fontFallbacks]);

  // biome-ignore lint/correctness/useExhaustiveDependencies: the revision is the sync trigger, compared by identity and never read - what the story now says is read off the host at sync time
  useLayoutEffect(() => {
    live.current?.sync();
  }, [revision]);

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
      className={[
        editorClassNames.noteRow,
        editing ? editorClassNames.noteRowOpen : null,
      ]
        .filter(Boolean)
        .join(" ")}
      role={name === undefined ? undefined : "group"}
      aria-label={name}
      aria-readonly={editing && readOnly ? true : undefined}
      style={hidden ? { visibility: "hidden" } : undefined}
      // The press is answered by the view the row is about to mount, which places the caret where
      // it landed; letting the browser select the markup first would leave that selection behind.
      // A press on the number is the way back to the text instead, which the view over an open row
      // answers for itself (`editor/notes/noteSurface`), so the two states do the same thing
      onMouseDown={
        editing || (!onPress && !onReturn)
          ? undefined
          : (event) => {
              event.preventDefault();
              if (onReturn && pressedTheNumber(event.target)) {
                onReturn();
                return;
              }
              onPress?.({ left: event.clientX, top: event.clientY });
            }
      }
    />
  );
});
