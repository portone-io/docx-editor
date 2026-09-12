/**
 * Which side story the caret is in, and what the one view over it is built from.
 *
 * One story is open at a time (`editor/stories/storyView`), and this is what holds that true: the
 * state here names one story key. Everything the view is built from arrives as one binding, so a
 * second kind of story is another binding rather than more of the component that renders this.
 * Where a story stands on the screen is the drawing side's (`./noteBands`).
 */

import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { NoteRow } from "../../editor/commands/noteQueries";
import { documentOf, type EditorDocument } from "../../editor/editorDocument";
import {
  type ActiveSurface,
  EVERY_CAPABILITY,
  type StoryCaret,
  type StoryExtensions,
  type StoryHost,
} from "../../editor/stories/storyView";
import type { StoryKey } from "../../schema/stories";
import type { RowEditing } from "./StoryRow";

/** A view and the state on screen, which is what the mounting component holds as React state */
export interface LiveSurface {
  readonly view: EditorView;
  readonly state: EditorState;
}

/** What one kind of story hands the view over one of its rows, which is all that differs per kind */
export interface StorySurfaceBinding {
  /** The host every edit in this kind of story leaves through */
  hostOf(main: EditorView, activate: StoryHost["activate"]): StoryHost;
  /** What the view over this row adds, and null for a row this binding does not answer for */
  extensionsOf(main: EditorView, row: NoteRow): StoryExtensions | null;
  /** The values this row's story is edited against, worked out from the document's own snapshot */
  documentFor(
    main: EditorState,
    snapshot: EditorDocument,
    row: NoteRow
  ): EditorDocument;
  /**
   * Which story the document is asking to have opened, read off the state the mounting component
   * holds. A fresh value for every request, so asking twice for one story opens it twice.
   */
  requestedIn(main: EditorState): { readonly key: StoryKey } | null;
  /** Puts the caret in this row's story, which is what opens it */
  openIn(main: EditorView, row: NoteRow): void;
  /** Hands the caret back to the text this row's story is called from, which is what closes it */
  returnFrom(main: EditorView, row: NoteRow): void;
}

export interface StorySurfaceOptions {
  /** The body's view and the state on screen, or null before a document is opened */
  readonly main: LiveSurface | null;
  /** The rows one of which a view may stand over, by the key naming each one's story */
  readonly rows: ReadonlyMap<StoryKey, NoteRow>;
  /** The kind of story these rows hold. A new value takes an open view down, so keep one */
  readonly binding: StorySurfaceBinding;
}

export interface StorySurface {
  /** Where the caret is, and null before a document is opened */
  readonly active: ActiveSurface | null;
  /** The story a view stands over, and null while the caret is in the body */
  readonly open: StoryKey | null;
  /** What mounts that view, which only the open row is given. Null while none is open */
  readonly editing: RowEditing | null;
  /** Whether the open story stands where nothing may be edited */
  readonly readOnly: boolean;
  /** Whether the open view is composing, which a page laid out again would take down */
  composing(): boolean;
  /** Opens the story of this row, with the caret where the press landed */
  onOpen(key: StoryKey, at: { left: number; top: number }): void;
  /** Hands the caret back to the text this row's story is called from */
  onReturn(key: StoryKey): void;
}

export function useStorySurface({
  main,
  rows,
  binding,
}: StorySurfaceOptions): StorySurface {
  const [asked, setAsked] = useState<StoryKey | null>(null);
  const [live, setLive] = useState<LiveSurface | null>(null);
  const held = useRef<{ key: StoryKey; caret: StoryCaret } | null>(null);

  // What asks for a story is the document's own state - a press on a note's number, a command, the
  // note just inserted - and the body taking the focus back, which is what Escape and a press on
  // the paper both end with, is what closes it
  const requested = main === null ? null : binding.requestedIn(main.state);
  useEffect(() => {
    if (requested !== null) setAsked(requested.key);
  }, [requested]);
  useEffect(() => {
    const body = main?.view;
    if (!body) return;
    const leave = () => setAsked(null);
    body.dom.addEventListener("focus", leave);
    return () => body.dom.removeEventListener("focus", leave);
  }, [main?.view]);

  // A story whose row an edit swept away has nowhere left to stand
  const open = asked !== null && rows.has(asked) ? asked : null;
  const row = open === null ? undefined : rows.get(open);

  const activate = useCallback((_key: StoryKey, view: EditorView | null) => {
    setLive(view === null ? null : { view, state: view.state });
  }, []);
  const onStateChange = useCallback((state: EditorState) => {
    setLive((current) =>
      current === null ? current : { view: current.view, state }
    );
  }, []);
  const caret = useMemo(
    () => ({
      take: () => (held.current?.key === open ? held.current.caret : null),
      keep: (at: StoryCaret) => {
        if (open !== null) held.current = { key: open, caret: at };
      },
    }),
    [open]
  );

  // biome-ignore lint/correctness/useExhaustiveDependencies: the host reads the state off the view on every call, so a state the body changed is no reason to build one again and take the open view down with it
  const host = useMemo(
    () => (main === null ? null : binding.hostOf(main.view, activate)),
    [main?.view, activate, binding]
  );
  // biome-ignore lint/correctness/useExhaustiveDependencies: the label is read through a closure the view asks as it draws, so it is a reason to build the extensions again and the row's identity is not
  const extensions = useMemo(
    () =>
      main === null || row === undefined
        ? null
        : binding.extensionsOf(main.view, row),
    [main?.view, row?.key, row?.label, binding]
  );
  const snapshot = main === null ? null : documentOf(main.state);
  // biome-ignore lint/correctness/useExhaustiveDependencies: a story is read against the styles the snapshot holds and the paper where its row is anchored, so an edit that moves neither must not build it again and take the open view down with it
  const document = useMemo(
    () =>
      main === null || snapshot === null || row === undefined
        ? null
        : binding.documentFor(main.state, snapshot, row),
    [main?.view, snapshot, row?.referencePos, binding]
  );

  const editing = useMemo<RowEditing | null>(
    () =>
      host === null || extensions === null || document === null
        ? null
        : { host, document, extensions, caret, onStateChange },
    [host, extensions, document, caret, onStateChange]
  );

  const active: ActiveSurface | null =
    main === null
      ? null
      : open !== null && live !== null && extensions !== null
        ? {
            surface: "story",
            key: open,
            view: live.view,
            state: live.state,
            takes: extensions.takes,
          }
        : {
            surface: "body",
            view: main.view,
            state: main.state,
            takes: EVERY_CAPABILITY,
          };

  return {
    active,
    open,
    editing,
    readOnly: open !== null && host !== null && host.shut(open),
    composing: () => live?.view.composing === true,
    onOpen(key, at) {
      const pressed = rows.get(key);
      if (main === null || pressed === undefined) return;
      held.current = { key, caret: { kind: "point", ...at } };
      binding.openIn(main.view, pressed);
    },
    onReturn(key) {
      const pressed = rows.get(key);
      if (main === null || pressed === undefined) return;
      // The body taking the focus back is what closes the story, open or not
      binding.returnFrom(main.view, pressed);
    },
  };
}
