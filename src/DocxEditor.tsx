/**
 * Renders one editing surface for a docx document, with the toolbar and the right click
 * menus around it.
 *
 * It keeps a second copy of the editor state that ProseMirror holds as React state,
 * so the buttons and the menus re-decide themselves every time the cursor moves.
 *
 * To swap the document, change the `key` and remount.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  type CSSProperties,
  type ForwardedRef,
  forwardRef,
  type ReactElement,
  type ReactNode,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { exportDocx } from "./docx/exportDocx";
import { type HeadersFooters, variantsFor } from "./docx/headersFooters";
import { type DocxBytes, type DocxSource, importDocx } from "./docx/importDocx";
import { type ExportProblem, exportProblems } from "./docx/invariants";
import { sectionIn, sectionsOf } from "./docx/sections";
import type { SessionStore } from "./docx/session";
import type { CommentAuthor } from "./editor/commands/commentCommands";
import { activeLinkSpan } from "./editor/commands/linkCommands";
import { createEditorView, editorStateForSession } from "./editor/createEditor";
import { sectionGeometryAt } from "./editor/documentStyles";
import { storyDocument } from "./editor/editorDocument";
import {
  noteExtensions,
  noteHost,
  returnToReference,
} from "./editor/notes/noteSurface";
import {
  closeCommentComposer,
  isCommentComposerOpen,
} from "./editor/plugins/commentComposer";
import { commentProjection } from "./editor/plugins/commentDecorations";
import { setProtection } from "./editor/plugins/documentProtection";
import { isLinkPanelOpen } from "./editor/plugins/linkPanel";
import {
  openNoteCommand,
  requestedNote,
} from "./editor/plugins/noteNavigation";
import { tableMenuAnchor } from "./editor/plugins/tableContextMenu";
import { textMenuAnchor } from "./editor/plugins/textContextMenu";
import { DocxImportError, type DocxImportErrorCode } from "./ooxml/errors";
import { PageGuides } from "./page/PageGuides";
import { A4_PAGE_PIXELS, pagePixels, sectionPixels } from "./page/pageLayout";
import { type PageFace, usePageLayout } from "./page/usePageLayout";
import type { EditableComments, EditingProtection } from "./schema/protection";
import { editingProtection, protectionOf } from "./schema/protectionState";
import { storyNodeOf } from "./schema/stories";
import { editorClassNames } from "./styles/classNames";
import type { FontFallbacks } from "./styles/fontStack";
import { CommentsPanel, shownBesideThePage } from "./ui/CommentsPanel";
import { LinkCard } from "./ui/LinkCard";
import { LinkPanel } from "./ui/LinkPanel";
import { NotesAroundPage, useNoteBands } from "./ui/notes/noteBands";
import {
  type StorySurfaceBinding,
  useStorySurface,
} from "./ui/notes/useStorySurface";
import type { DocxEditorPresets } from "./ui/presets";
import { TableMenu } from "./ui/TableMenu";
import { TextMenu } from "./ui/TextMenu";
import { Toolbar } from "./ui/Toolbar";
import { useFitWidthZoom } from "./ui/useFitWidthZoom";
import { usePageRoom } from "./ui/usePageRoom";
import { type DocxEditorZoom, normalizeZoom } from "./ui/zoom";

export interface DocxEditorHandle {
  view: EditorView;
  /** Turns the editor state currently on screen into docx bytes */
  exportBytes: () => Uint8Array;
  /**
   * Every reason `exportBytes` would refuse the state currently on screen, each under the code
   * the refusal would carry; empty when it would write. `downloadDocx` answers `blocked` with it
   */
  exportProblems: () => readonly ExportProblem[];
}

/**
 * What the editor is for, which decides what it offers.
 *
 * The three kinds are the standings OOXML document protection names (`ST_DocProtect`) and the
 * ones a shared document gives its readers: a reader, a commenter, an editor. A read-only editor
 * takes no edits, so it has no toolbar and no right click menus either: what used to be three
 * booleans of which one silently emptied the other two is one choice here. A `comment` editor is
 * the `comments` protection - the body may not be changed, comments may be written, answered and
 * settled - which is what a reviewer who must not touch the text is handed. Both kinds that write
 * comments name whose they are, so a reader is never asked for an identity it has no use for.
 *
 * `editableComments` is whose comments the panel offers to edit or delete: one's own, which is a
 * comment carrying no recognised identity or the very identity given here, or every one, which
 * is a moderator's standing. Replying and resolving are open to every commenter either way.
 *
 * `locking` is what a screen where a template is written gets: settling a part of a document is
 * an authoring act, not something every reader of a form should be handed. A lock the document
 * already carries holds whichever mode is chosen.
 *
 * The kind, the author and `editableComments` are read on every render and take effect on the
 * open document. Whether the right click opens the editor's own menus is `contextMenus`, a prop
 * of its own, since the plugins behind it are read when the editor mounts rather than per render.
 */
export type DocxEditorMode =
  | { kind: "readOnly" }
  | {
      kind: "comment";
      author: CommentAuthor;
      editableComments?: EditableComments;
    }
  | {
      kind: "edit";
      author: CommentAuthor;
      editableComments?: EditableComments;
      toolbar?: boolean;
      locking?: boolean;
    };

/**
 * Runs a command on the view without taking the focus, which stays with whatever control asked.
 * The runner in `ui/runCommand` hands the focus back to the text, which is what a control acting
 * on the text wants; a control acting on the screen around it does not.
 */
function runOn(view: EditorView, command: Command): void {
  command(view.state, (transaction) => view.dispatch(transaction));
}

/**
 * What the note surface hands the view over one note, which is the whole of what this component
 * knows about a kind of story (`editor/notes/noteSurface`).
 *
 * One binding answers for a footnote and an endnote alike: where each is drawn is the drawing
 * side's question, and what an edit in one does is the same.
 */
const NOTE_SURFACE: StorySurfaceBinding = {
  hostOf: noteHost,
  extensionsOf: (main, row) => noteExtensions(main, row.key, () => row.label),
  // The paper a note wraps at is the paper of the section its reference stands in, which a story
  // holds no section of its own to say (`editor/documentStyles`)
  documentFor: (main, snapshot, row) =>
    storyDocument(snapshot, sectionGeometryAt(main, row.referencePos)),
  requestedIn: requestedNote,
  // The kind's own public command is what a consumer calls; here the kind is a value the row
  // carries, so the factory behind the two takes it
  openIn: (main, row) => runOn(main, openNoteCommand(row.kind, row.id)),
  returnFrom: (main, row) => returnToReference(main, row.key),
};

/** What a mode hands the reader, which is everything the component reads off the kind */
interface ModeAffordances {
  /** The protection (`schema/protection`) the editor state is put under */
  protection: EditingProtection;
  /** The identity comments are written under. Null for a reader, who writes none */
  author: CommentAuthor | null;
  editableComments: EditableComments;
  toolbar: boolean;
  locking: boolean;
}

/**
 * Every kind answered in one exhaustive place, so that a kind added later - a suggester, a form to
 * fill in - is a compile error here rather than a silent `false` at each `kind === "edit"` the
 * component would otherwise ask.
 *
 * An author is read through `?? null` because a consumer writing no TypeScript can leave it out,
 * and a composer offered to nobody is a better answer than one that reads a name off nothing.
 */
function affordancesOf(mode: DocxEditorMode): ModeAffordances {
  switch (mode.kind) {
    case "readOnly":
      return {
        protection: "readOnly",
        author: null,
        editableComments: "own",
        toolbar: false,
        locking: false,
      };
    case "comment":
      return {
        protection: "comments",
        author: mode.author ?? null,
        editableComments: mode.editableComments ?? "own",
        toolbar: false,
        locking: false,
      };
    case "edit":
      return {
        protection: "none",
        author: mode.author ?? null,
        editableComments: mode.editableComments ?? "own",
        toolbar: mode.toolbar ?? true,
        locking: mode.locking ?? false,
      };
    default: {
      const unmodelled: never = mode;
      return unmodelled;
    }
  }
}

export interface DocxEditorProps {
  /**
   * The document to open: the bytes, or the `File`/`Blob` a file input or a fetch
   * hands over, which the editor reads itself.
   *
   * Bytes open on the spot. A `Blob` is read first, so the editor stands empty for
   * that moment; whichever document arrived last is the one opened.
   */
  document: DocxSource;
  /**
   * What to render instead when the document could not be opened.
   *
   * A built-in panel naming the reason is drawn when none is given, so a refusal is
   * never silent either way. Hand one in to write the refusal in your own words.
   */
  renderImportError?: (error: DocxImportError) => ReactNode;
  /** What the editor is for: a reader's, a commenter's or an editor's surface */
  mode: DocxEditorMode;
  /** Whether to draw approximate page boundaries over the document's own paper. Drawn when read only too */
  showPageGuides?: boolean;
  /**
   * The document's visual scale. `fit-width` follows the editor's available width;
   * a number uses that fixed scale without changing document layout or export.
   */
  zoom?: DocxEditorZoom;
  /** Read once for uncontrolled state. Later changes are ignored. Defaults to `fit-width` */
  defaultZoom?: DocxEditorZoom;
  /** Receives toolbar requests; a controlled consumer must update `zoom` to apply one */
  onZoomChange?: (zoom: DocxEditorZoom) => void;
  /**
   * The fonts drawn in place of the ones a document declares, for the names that
   * are missing from the reader's machine.
   * The built-in set knows the CJK and Latin office font names and lands on a
   * Latin sans for the rest; hand in your own set to stand in different fonts.
   * A set handed in reaches the runs and the paper, while paragraph styles and
   * the HTML copied out keep the built-in set, which is shared beyond one editor.
   * It only ever affects what is drawn - the exported document keeps the fonts it
   * declared.
   *
   * Read once, when the editor mounts, like `plugins`. Later changes to it are
   * ignored, so writing the set inline on every render is harmless; to change the
   * fallbacks, change the `key` and remount.
   */
  fontFallbacks?: FontFallbacks;
  /**
   * The lists the toolbar's pickers offer: fonts, colors, font sizes, line spacings and
   * cell borders.
   * Every list left out is the built-in one. A font or a size the open document itself
   * uses is offered whether or not the list handed in holds it.
   *
   * Read on every render, unlike `fontFallbacks` and `plugins`, so a list built from state
   * takes effect without a remount.
   */
  presets?: DocxEditorPresets;
  /**
   * ProseMirror plugins handed in by the consumer, which is how mentions, highlights,
   * autocomplete and the like are added from outside the package.
   * They are placed ahead of the built-in plugins, so a consumer keymap or DOM handler
   * sees an event first and can override a built-in shortcut.
   * A read-only editor gets them too.
   *
   * Read once, when the editor mounts. Later changes to this array are ignored, so
   * building it inline on every render is harmless; to swap the plugins, change the
   * `key` and remount.
   */
  plugins?: readonly Plugin[];
  /**
   * Whether the right click opens the editor's own menus. `false` leaves it to the browser,
   * which is what a consumer drawing menus of its own wants. Defaults to `true`.
   *
   * Read once, when the editor mounts, like `plugins`: the plugins that take the browser's menu
   * away go into the editor state, which is built there. A menu the protection has nothing to
   * offer from stands down by itself, so a mode change needs nothing here. Later changes are
   * ignored; to turn the menus around, change the `key` and remount.
   */
  contextMenus?: boolean;
  className?: string;
  style?: CSSProperties;
  onReady?: (view: EditorView) => void;
  /** Called every time the editor state changes. This covers cursor and selection moves, not just text edits */
  onChange?: () => void;
}

type OpenedDocument =
  | {
      status: "opened";
      doc: PMNode;
      session: SessionStore;
    }
  | { status: "rejected"; error: DocxImportError };

interface LiveEditor {
  view: EditorView;
  state: EditorState;
}

/**
 * Only a rejection for want of a preservation guarantee comes back as state; every
 * other error is rethrown as is.
 */
function openDocument(bytes: DocxBytes): OpenedDocument {
  try {
    const { doc, session } = importDocx(bytes);
    return { status: "opened", doc, session };
  } catch (error) {
    if (error instanceof DocxImportError) return { status: "rejected", error };
    throw error;
  }
}

/** Callbacks are kept in a box so that a parent re-render does not rebuild the editor */
function useLatest<T>(value: T) {
  const box = useRef(value);
  useEffect(() => {
    box.current = value;
  });
  return box;
}

function isBlob(source: DocxSource): source is Blob {
  return typeof Blob !== "undefined" && source instanceof Blob;
}

/**
 * What reading a `Blob` left behind, tagged with the blob it was read from.
 *
 * The tag is what makes a stale read harmless: a document swapped mid-read is a
 * different blob, so the bytes of the one before it are never opened even on the
 * render between the swap and the next read.
 */
type BlobRead =
  | { of: Blob; status: "read"; bytes: ArrayBuffer }
  | { of: Blob; status: "failed"; error: unknown };

/**
 * The bytes to open, or null while a blob is still being read.
 * Bytes handed in directly are returned as they are, so nothing about the
 * synchronous path changes.
 */
function useDocumentBytes(source: DocxSource): DocxBytes | null {
  const [read, setRead] = useState<BlobRead | null>(null);

  useEffect(() => {
    if (!isBlob(source)) return;
    let current = true;
    source
      .arrayBuffer()
      .then((bytes) => {
        if (current) setRead({ of: source, status: "read", bytes });
      })
      .catch((error: unknown) => {
        if (current) setRead({ of: source, status: "failed", error });
      });
    return () => {
      current = false;
    };
  }, [source]);

  if (!isBlob(source)) return source;
  if (read?.of !== source) return null;
  // A blob that cannot be read is not a document we refused; it never arrived
  if (read.status === "failed") throw read.error;
  return read.bytes;
}

const IMPORT_REJECTION_REASON: Record<DocxImportErrorCode, string> = {
  // `DOMParser` is built into the browser rather than loaded with the page, so a browser that
  // reaches this line has no XML reader at all and will not have gained one on a second try
  "no-xml-parser": "This browser cannot read documents.",
  "not-a-docx": "This file is not a docx document, or it is damaged.",
  "too-large": "This document is too large to open.",
  "missing-part": "This document is missing the part that holds its body.",
  "missing-body": "This document has no body.",
  "malformed-xml": "The XML inside this document cannot be read.",
  "unsupported-conformance":
    "This document is saved in a format this editor cannot open. Save it as a Word document and try again.",
  "unsupported-content":
    "This document holds content that could not be kept as it is.",
};

/**
 * What a refusal looks like when the consumer wrote no panel of its own.
 *
 * The document is not opened either way; this is what makes that visible rather
 * than leaving an empty box behind.
 */
function ImportRejection({ error }: { error: DocxImportError }): ReactElement {
  return (
    <div className={editorClassNames.rejection} role="alert">
      <p className={editorClassNames.rejectionTitle}>
        This document was not opened
      </p>
      <p>{IMPORT_REJECTION_REASON[error.code]}</p>
    </div>
  );
}

/**
 * Publishes the zoom factor to CSS. It is what scales the page layer itself
 * (`styles/editor.css`), and it reaches the panels beside the paper, which stand outside
 * the layer and scale their own type with it.
 */
function zoomVariable(
  factor: number
): CSSProperties & Record<"--docx-editor-zoom", number> {
  return { "--docx-editor-zoom": factor };
}

function DocxEditorSurface(
  {
    document: source,
    renderImportError,
    mode,
    showPageGuides = true,
    zoom,
    defaultZoom = "fit-width",
    onZoomChange,
    fontFallbacks,
    presets,
    plugins,
    contextMenus,
    className,
    style,
    onReady,
    onChange,
  }: DocxEditorProps,
  ref: ForwardedRef<DocxEditorHandle | null>
): ReactNode {
  const { protection, author, editableComments, toolbar, locking } =
    affordancesOf(mode);
  const bytes = useDocumentBytes(source);
  const opened = useMemo(
    () => (bytes === null ? null : openDocument(bytes)),
    [bytes]
  );
  // Held from the first render on, so a value rebuilt on every render does not rebuild the editor
  const mountedPlugins = useRef(plugins).current;
  // The plugins are mounted whatever the mode is, and stand down while the protection shuts what
  // they offer, so that opening the document up again brings the menus back with it
  const mountedContextMenus = useRef(contextMenus ?? true).current;
  const mountedFontFallbacks = useRef(fontFallbacks).current;
  const layerRef = useRef<HTMLDivElement | null>(null);
  const boxRef = useRef<HTMLDivElement | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const mountRef = useRef<HTMLDivElement | null>(null);
  const keptState = useRef<{ of: OpenedDocument; state: EditorState } | null>(
    null
  );
  const viewRef = useRef<EditorView | null>(null);
  const [live, setLive] = useState<LiveEditor | null>(null);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [uncontrolledZoom, setUncontrolledZoom] = useState<DocxEditorZoom>(() =>
    normalizeZoom(defaultZoom)
  );
  const latestOnReady = useLatest(onReady);
  const latestOnChange = useLatest(onChange);
  const selectedZoom = normalizeZoom(zoom ?? uncontrolledZoom);
  // One sheet is drawn at one paper, the first section's, whatever the sections after it name
  // (`page/pageLayout`)
  const page =
    opened?.status === "opened"
      ? pagePixels(opened.session.geometry)
      : A4_PAGE_PIXELS;
  const fitWidth = useFitWidthZoom(rootRef, page.pageWidth);
  const effectiveZoom = selectedZoom === "fit-width" ? fitWidth : selectedZoom;
  usePageRoom(boxRef, layerRef, effectiveZoom);
  const changeZoom = (next: DocxEditorZoom) => {
    const normalized = normalizeZoom(next);
    if (zoom === undefined) setUncontrolledZoom(normalized);
    onZoomChange?.(normalized);
  };

  // Mounted right after render, so a consumer can read the ref straight from its own effect.
  // biome-ignore lint/correctness/useExhaustiveDependencies: the protection is read when the state is first built; a later mode goes in through the effect below rather than through a new view
  useLayoutEffect(() => {
    const mount = mountRef.current;
    if (!mount || opened?.status !== "opened") return;

    const kept = keptState.current;
    const view = createEditorView({
      mount,
      state:
        kept?.of === opened
          ? kept.state
          : editorStateForSession(opened, {
              consumerPlugins: mountedPlugins,
              contextMenus: mountedContextMenus,
              protection,
              author,
              editableComments,
            }),
      fontFallbacks: mountedFontFallbacks,
      onStateChange: (state) => {
        keptState.current = { of: opened, state };
        setLive({ view, state });
        latestOnChange.current?.();
      },
    });
    viewRef.current = view;
    keptState.current = { of: opened, state: view.state };
    setLive({ view, state: view.state });
    latestOnReady.current?.(view);

    return () => {
      view.destroy();
      viewRef.current = null;
      setLive(null);
    };
  }, [
    opened,
    mountedContextMenus,
    mountedFontFallbacks,
    mountedPlugins,
    latestOnReady,
    latestOnChange,
  ]);

  // A mode changed on an open document is put into the state it already holds, so the view, its
  // history and the consumer's plugins all stay: `reconfigure` would keep the protection plugin's
  // state too, which is why it goes in as a transaction.
  //
  // It goes in from a layout effect, so the dispatch and the render it leaves behind both close
  // before the browser paints: no frame is drawn with the state under one protection and the
  // controls around it under another.
  const authorId = author?.id ?? null;
  const mounted = live !== null;
  // biome-ignore lint/correctness/useExhaustiveDependencies: a mode written inline is a new object on every render, so the author is watched through `authorId`, the part of it the protection holds
  useLayoutEffect(() => {
    const view = viewRef.current;
    if (!view || !mounted) return;
    const held = protectionOf(view.state);
    if (
      held.protection === protection &&
      held.authorId === authorId &&
      held.editableComments === editableComments
    ) {
      return;
    }
    view.dispatch(
      setProtection(view.state.tr, { protection, author, editableComments })
    );
  }, [mounted, protection, authorId, editableComments]);

  useImperativeHandle<
    DocxEditorHandle | null,
    DocxEditorHandle | null
  >(ref, () => {
    const view = viewRef.current;
    if (!view || opened?.status !== "opened") return null;
    const session = opened.session;
    return {
      view,
      exportBytes: () => exportDocx(view.state.doc, session),
      exportProblems: () => exportProblems(view.state.doc, session),
    };
  }, [opened]);

  // The section table is read once per document rather than once per page or per measurement:
  // reading it walks every block, and both the page arithmetic and the preview ask about it for
  // each page they draw
  const doc = live?.state.doc ?? null;
  const session = opened?.status === "opened" ? opened.session : null;
  const sections = useMemo(
    () => (doc === null ? null : sectionsOf(doc)),
    [doc]
  );
  const sectionPapers = useMemo(
    () => (sections === null ? undefined : sectionPixels(sections)),
    [sections]
  );

  // What the notes ask a page to keep at its foot, and what draws them, in one place
  // (`ui/notes/noteBands`): a second band is another of these rather than more of this component
  const notes = useNoteBands({
    state: live?.state ?? null,
    of: opened,
    fontFallbacks: mountedFontFallbacks,
  });

  // Which note the caret is in and what the view over it is built from, in one place
  // (`ui/notes/useStorySurface`): a second kind of story is another binding rather than more of
  // this component
  const surface = useStorySurface({
    main: live,
    rows: notes.rows,
    binding: NOTE_SURFACE,
  });

  const overlay = usePageLayout({
    view: live?.view ?? null,
    layer: layerRef,
    enabled: showPageGuides,
    revision: live?.state.doc,
    sections: sectionPapers,
    bands: notes.bands,
    trailing: notes.trailing,
    // A page laid out again draws the row the open note stands in wherever it now lands, which
    // would take its composition down with it
    composing: surface.composing,
  });

  const headersFootersFor = useMemo(() => {
    if (doc === null || session === null || sections === null) return undefined;
    const shown = sections.map((section) =>
      variantsFor(section, session.headerFooterStories, (key) =>
        storyNodeOf(doc, key)
      )
    );
    return (face: PageFace): HeadersFooters | null =>
      shown[sectionIn(sections, doc, face.pos).index] ?? null;
  }, [doc, session, sections]);

  if (opened?.status === "rejected") {
    return renderImportError ? (
      renderImportError(opened.error)
    ) : (
      <ImportRejection error={opened.error} />
    );
  }

  // What the state on screen lets through, which is what every control below asks rather than
  // the mode itself. The protection goes in from a layout effect above, so the state a control
  // reads and the mode it was drawn for are the same on every frame the reader sees
  const bodyOpen = live !== null && editingProtection(live.state) === "none";

  // Only ever one menu at a time. Where both plugins hold a point, the text menu is the one the
  // last right click opened, so it is the one drawn.
  // Where the menu plugins were left out, no point is ever held and neither menu is drawn.
  const textAnchor = live ? textMenuAnchor(live.state) : null;
  const tableAnchor = live && !textAnchor ? tableMenuAnchor(live.state) : null;

  // The card that says where a link points stands down while the panel that changes one is open, and
  // while an IME is composing: a box appearing and moving under a composition is the churn the page
  // measurement documents avoiding (`page/usePageLayout`)
  const linkAtCursor =
    live && !live.view.composing && !isLinkPanelOpen(live.state)
      ? activeLinkSpan(live.state)
      : null;
  const comments =
    live?.state === undefined
      ? []
      : commentProjection.read(live.state).comments;
  const hasRailComments = comments.some(shownBesideThePage);
  // The composer's own state closes it where a comment can no longer go, the mode being switched
  // to read-only included, so nothing here has to shut it in turn
  const composerOpen = live !== null && isCommentComposerOpen(live.state);
  const showComments =
    live !== null && (commentsOpen || composerOpen || hasRailComments);
  const commentsPanel = live && showComments && (
    <CommentsPanel
      view={live.view}
      state={live.state}
      author={author}
      scrollContainer={rootRef.current}
      allCommentsOpen={commentsOpen}
    />
  );

  return (
    <div
      className={[editorClassNames.frame, className].filter(Boolean).join(" ")}
      style={style}
    >
      {live && toolbar && surface.active && (
        <Toolbar
          main={live}
          active={surface.active}
          fontFallbacks={mountedFontFallbacks}
          presets={presets}
          commentsOpen={commentsOpen}
          onToggleComments={() => {
            // Putting the list away puts away the form standing in it
            if (commentsOpen) runOn(live.view, closeCommentComposer);
            setCommentsOpen(!commentsOpen);
          }}
          zoom={selectedZoom}
          onZoomChange={changeZoom}
        />
      )}
      <div
        className={editorClassNames.workspace}
        data-comments={showComments ? "visible" : undefined}
        style={zoomVariable(effectiveZoom)}
      >
        {/* Where the body is shut there is no toolbar, so the comments are reached from here */}
        {live && !bodyOpen && comments.length > 0 && (
          <button
            type="button"
            className={editorClassNames.commentsToggle}
            aria-expanded={commentsOpen}
            onClick={() => setCommentsOpen((open) => !open)}
          >
            {commentsOpen ? "Hide comments" : "Show comments"}
          </button>
        )}
        <div ref={rootRef} className={editorClassNames.root}>
          {/* The scaled layer stands outside the flow, so this box holds the room it takes */}
          <div ref={boxRef} className={editorClassNames.pageBox}>
            {/* The paper and the page marks overlaid on it share one positioning origin */}
            <div ref={layerRef} className={editorClassNames.pageLayer}>
              <div ref={mountRef} />
              {overlay && (
                <PageGuides
                  overlay={overlay}
                  headersFootersFor={headersFootersFor}
                />
              )}
              <NotesAroundPage
                notes={notes}
                overlay={overlay}
                page={page}
                zoom={effectiveZoom}
                open={surface.open}
                editing={surface.editing}
                readOnly={surface.readOnly}
                revision={live?.state}
                onOpen={surface.onOpen}
                onReturn={surface.onReturn}
              />
            </div>
          </div>
          {!commentsOpen && commentsPanel}
        </div>
        {commentsOpen && commentsPanel}
      </div>
      {live && bodyOpen && isLinkPanelOpen(live.state) && (
        <LinkPanel view={live.view} state={live.state} />
      )}
      {live && linkAtCursor && (
        <LinkCard
          // Escape hides the card over the link it was pressed on; another link is another card
          key={linkAtCursor.from}
          view={live.view}
          state={live.state}
          link={linkAtCursor}
        />
      )}
      {live && textAnchor && surface.active && (
        <TextMenu
          view={live.view}
          state={live.state}
          anchor={textAnchor}
          takes={surface.active.takes}
          allowLocking={locking}
        />
      )}
      {live && tableAnchor && surface.active && (
        <TableMenu
          view={live.view}
          state={live.state}
          anchor={tableAnchor}
          takes={surface.active.takes}
          allowLocking={locking}
        />
      )}
    </div>
  );
}

/**
 * `forwardRef` rather than a `ref` prop: React 18 hands a function component no `ref`
 * in its props, so the handle would never be built there.
 */
export const DocxEditor = forwardRef(DocxEditorSurface);
