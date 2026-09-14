/**
 * The one editing view a side story is edited in, mounted over the story the caret went into and
 * taken down when it leaves.
 *
 * Everything typed here leaves through the host, which writes it into the main document as one
 * story change. That is what keeps the guards, the locks and the protection judging a story edit
 * where they judge a body edit, and what leaves the main document holding the only history: the
 * undo keys here run the host's.
 *
 * This file does not know which kind of story it is drawing or where it is drawn: a note stands
 * at the foot of a page, a comment body in a rail beside it, a header in a page margin. What
 * differs between the kinds arrives as `StoryExtensions`.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  type Plugin,
  TextSelection,
} from "prosemirror-state";
import { EditorView, type NodeViewConstructor } from "prosemirror-view";
import type { EditingProtection } from "../../schema/protection";
import { editingProtection, editsShut } from "../../schema/protectionState";
import { type StoryKey, sameStory, storyNodeOf } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import type { FontFallbacks } from "../../styles/fontStack";
import type { SliceNormalizer } from "../clipboard/normalizers";
import type { EditorDocument } from "../editorDocument";
import { runMarkView } from "../views/runMarkView";
import { storyEditorState } from "./storyState";

export type NodeViewMap = Record<string, NodeViewConstructor>;

/** What every surface is read and dispatched through, whichever of the two it is */
interface SurfaceContents {
  readonly view: EditorView;
  /** The state on screen, which a control drawn from the surface re-decides itself from */
  readonly state: EditorState;
  readonly takes: SurfaceCapabilities;
}

/**
 * The surface holding the caret, which is either the body or one story with the view over it.
 *
 * Two editing views cannot stand at once, and this is where that rule is written down: a caller
 * holding this value has no way to name a second one. What holds it true is the state that decides
 * which story is open (`ui/notes/useStorySurface`).
 */
export type ActiveSurface =
  | ({ readonly surface: "body" } & SurfaceContents)
  | ({ readonly surface: "story"; readonly key: StoryKey } & SurfaceContents);

export interface StoryHost {
  /** The main state the story is read from and judged against */
  readonly state: () => EditorState;
  /** Writes the story into the main document; false when the main document refused it */
  write(key: StoryKey, story: PMNode): boolean;
  undo(): boolean;
  redo(): boolean;
  /** Hands the caret back to the main document (Escape) */
  leave(key: StoryKey): void;
  /** Whether edits to this story are closed where it is anchored */
  shut(key: StoryKey): boolean;
  /** Registers the view that holds the caret, null when it lets go */
  activate(key: StoryKey, view: EditorView | null): void;
}

/**
 * What a surface takes beyond character and paragraph formatting.
 *
 * Every one of these puts something into the document that the part the surface is written back as
 * has to carry - a relationship, a definition, a part of its own - so which of them a surface takes
 * is the surface's own to declare, and a control or a key that puts one in asks here rather than
 * asking which kind of surface it is.
 */
export const SURFACE_CAPABILITIES = [
  "list",
  "table",
  "image",
  "link",
  "note",
  "comment",
] as const;

export type SurfaceCapability = (typeof SURFACE_CAPABILITIES)[number];

export type SurfaceCapabilities = ReadonlySet<SurfaceCapability>;

/** What the body takes, which is everything: it is written back as the document part itself */
export const EVERY_CAPABILITY: SurfaceCapabilities = new Set(
  SURFACE_CAPABILITIES
);

/** What a surface whose writer carries none of them takes */
export const NO_CAPABILITY: SurfaceCapabilities = new Set();

export interface StoryExtensions {
  readonly plugins: readonly Plugin[];
  readonly nodeViews: NodeViewMap;
  readonly normalizers: readonly SliceNormalizer[];
  /** What this kind of story takes, which its keys and the controls over it are drawn from */
  readonly takes: SurfaceCapabilities;
}

/**
 * Where the caret goes when the view is mounted: back where it stood, or at the point the reader
 * pressed, which is what a press on a story no view stands over yet has to say.
 */
export type StoryCaret =
  | { readonly kind: "at"; readonly anchor: number; readonly head: number }
  | { readonly kind: "point"; readonly left: number; readonly top: number };

export interface StoryViewOptions {
  readonly mount: HTMLElement;
  readonly host: StoryHost;
  readonly key: StoryKey;
  readonly document: EditorDocument;
  readonly fontFallbacks: FontFallbacks;
  readonly extensions: StoryExtensions;
  /** Where the caret opens. The end of the story when none is given */
  readonly caret?: StoryCaret | null;
  /**
   * Called with every state this view goes to, which is what a control drawn from the story - a
   * toolbar acting on it - re-decides itself from.
   */
  readonly onStateChange?: (state: EditorState) => void;
}

export interface StoryView {
  readonly view: EditorView;
  /** Takes a changed main state in; replaces the content only when the story says something else */
  sync(): void;
  destroy(): void;
}

/** Where the caret stands, in the shape a view mounted again over the same story opens at */
export function caretOf(state: EditorState): StoryCaret {
  return {
    kind: "at",
    anchor: state.selection.anchor,
    head: state.selection.head,
  };
}

function placed(state: EditorState, at: number, to: number): EditorState {
  const size = state.doc.content.size;
  const $anchor = state.doc.resolve(Math.min(Math.max(at, 0), size));
  const $head = state.doc.resolve(Math.min(Math.max(to, 0), size));
  return state.apply(
    state.tr.setSelection(TextSelection.between($anchor, $head))
  );
}

/**
 * Ends the composition a refusal has already broken.
 *
 * Handing ProseMirror stored marks is the door it ends one through of its own accord, which is
 * how the body's own refusals are answered (`editor/plugins/lockedContent`).
 */
function endComposition(view: EditorView): void {
  if (!view.composing) return;
  const marks = view.state.storedMarks ?? view.state.selection.$from.marks();
  view.dispatch(
    view.state.tr.setStoredMarks(marks).setMeta("addToHistory", false)
  );
}

export function createStoryView({
  mount,
  host,
  key,
  document,
  fontFallbacks,
  extensions,
  caret = null,
  onStateChange,
}: StoryViewOptions): StoryView {
  /**
   * What the story runs under: the main document's standing, and nothing at all where the place
   * the story is called from is shut. A lock around the reference therefore shuts the story the
   * same way it shuts the text around it.
   */
  const standing = (): EditingProtection =>
    host.shut(key) ? "readOnly" : editingProtection(host.state());

  let protection = standing();

  const keys: Record<string, Command> = {
    "Mod-z": () => host.undo(),
    "Mod-y": () => host.redo(),
    "Shift-Mod-z": () => host.redo(),
    Escape: () => {
      host.leave(key);
      return true;
    },
  };

  const built = (story: PMNode): EditorState =>
    storyEditorState({
      story,
      document,
      protection,
      keys,
      plugins: extensions.plugins,
      normalizers: extensions.normalizers,
      takes: extensions.takes,
    });

  const storyNow = (): PMNode | null => storyNodeOf(host.state().doc, key);

  const opening = storyNow();
  if (opening === null) {
    throw new Error(`the document holds no story named ${key}`);
  }

  const view = new EditorView(mount, {
    state: built(opening),
    // A standing that shuts the story shuts typing in it, and is read off the state so that a
    // mode switched while the story is open takes effect without a new view
    editable: (current) => !editsShut(current),
    attributes: { class: editorClassNames.storyBody },
    // The schema can only draw a run with the default fallbacks; this view draws the host's
    markViews: { run: runMarkView(fontFallbacks) },
    nodeViews: extensions.nodeViews,
    dispatchTransaction(transaction) {
      const next = view.state.applyTransaction(transaction);
      if (next.state.doc === view.state.doc) {
        show(next.state);
        return;
      }
      // The edit is shown before it is offered, so that a host writing it back on the spot finds
      // the view already saying what it wrote and leaves the composition standing
      const before = view.state;
      show(next.state);
      if (host.write(key, next.state.doc)) return;
      // A host with nothing to write is not a host turning the edit down: an edit that leaves the
      // story saying what it already said - the same words pasted over themselves - is written
      // nowhere, and rewinding it would break the composition it stands in
      if (sameStory(storyNow(), next.state.doc)) return;
      show(before);
      endComposition(view);
    },
  });

  const show = (next: EditorState): void => {
    view.updateState(next);
    onStateChange?.(next);
  };

  const open = caret ?? null;
  if (open === null) {
    const end = view.state.doc.content.size;
    show(placed(view.state, end, end));
  } else if (open.kind === "at") {
    show(placed(view.state, open.anchor, open.head));
  } else {
    const at = view.posAtCoords({ left: open.left, top: open.top });
    const pos = at?.pos ?? view.state.doc.content.size;
    show(placed(view.state, pos, pos));
  }

  const registerFocus = () => host.activate(key, view);
  view.dom.addEventListener("focus", registerFocus);

  return {
    view,
    sync() {
      const story = storyNow();
      // A story the main document no longer holds is one whose anchor an edit took away; what
      // draws this view is what takes it down
      if (story === null) return;
      const wanted = standing();
      if (wanted === protection && sameStory(story, view.state.doc)) return;
      const { anchor, head } = view.state.selection;
      protection = wanted;
      show(placed(built(story), anchor, head));
    },
    destroy() {
      view.dom.removeEventListener("focus", registerFocus);
      host.activate(key, null);
      view.destroy();
    },
  };
}
