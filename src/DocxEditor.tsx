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
import type { EditorState, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  type CSSProperties,
  type ReactNode,
  type Ref,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { exportDocx } from "./docx/exportDocx";
import { type DocxBytes, type DocxSource, importDocx } from "./docx/importDocx";
import type { SessionStore } from "./docx/session";
import {
  type CommentAuthor,
  documentComments,
} from "./editor/commands/commentCommands";
import { activeLinkSpan } from "./editor/commands/linkCommands";
import { createEditorState, createEditorView } from "./editor/createEditor";
import { isLinkPanelOpen } from "./editor/plugins/linkPanel";
import { tableMenuAnchor } from "./editor/plugins/tableContextMenu";
import { textMenuAnchor } from "./editor/plugins/textContextMenu";
import { type Numbering, parseNumbering } from "./numbering/parseNumbering";
import { DocxImportError, type DocxImportErrorCode } from "./ooxml/errors";
import { PageGuides } from "./page/PageGuides";
import { A4_PAGE_PIXELS, pagePixels } from "./page/pageLayout";
import { usePageLayout } from "./page/usePageLayout";
import { editorClassNames } from "./styles/classNames";
import type { FontFallbacks } from "./styles/fontStack";
import { CommentsPanel } from "./ui/CommentsPanel";
import { LinkCard } from "./ui/LinkCard";
import { LinkPanel } from "./ui/LinkPanel";
import { NotesPanel } from "./ui/NotesPanel";
import type { DocxEditorPresets } from "./ui/presets";
import { TableMenu } from "./ui/TableMenu";
import { TextMenu } from "./ui/TextMenu";
import { Toolbar } from "./ui/Toolbar";
import { useFitWidthZoom } from "./ui/useFitWidthZoom";
import { type DocxEditorZoom, normalizeZoom } from "./ui/zoom";

export interface DocxEditorHandle {
  view: EditorView;
  /** Turns the editor state currently on screen into docx bytes */
  exportBytes: () => Uint8Array;
}

/**
 * What the editor is for, which decides what it offers.
 *
 * A read-only editor takes no edits, so it has no toolbar and no right click menus either: what
 * used to be three booleans of which one silently emptied the other two is one choice here.
 * `locking` is what a screen where a template is written gets: settling a part of a document is
 * an authoring act, not something every reader of a form should be handed. A lock the document
 * already carries holds whichever mode is chosen.
 *
 * `contextMenus: false` leaves the right click to the browser, for a consumer drawing menus of
 * its own. The plugins that take the browser's menu away go into the editor state, which is built
 * when the editor mounts, so this is read once there like `plugins`.
 */
export type DocxEditorMode =
  | { kind: "readOnly" }
  | {
      kind: "edit";
      toolbar?: boolean;
      contextMenus?: boolean;
      locking?: boolean;
    };

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
  ref?: Ref<DocxEditorHandle | null>;
  /** What the editor is for. Editing with the toolbar shown and no locking when none is given */
  mode?: DocxEditorMode;
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
  /** The identity written by the built-in comment and reply composers */
  commentAuthor?: CommentAuthor;
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
      numbering: Numbering;
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
    return {
      status: "opened",
      doc,
      session,
      numbering: parseNumbering(session.numberingXml),
    };
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
  "not-a-docx": "This file is not a docx document, or it is damaged.",
  "too-large": "This document is too large to open.",
  "missing-part": "This document is missing the part that holds its body.",
  "missing-body": "This document has no body.",
  "malformed-xml": "The XML inside this document cannot be read.",
  "unsupported-content":
    "This document holds content that could not be kept as it is.",
};

/**
 * What a refusal looks like when the consumer wrote no panel of its own.
 *
 * The document is not opened either way; this is what makes that visible rather
 * than leaving an empty box behind.
 */
function ImportRejection({ error }: { error: DocxImportError }) {
  return (
    <div className={editorClassNames.rejection} role="alert">
      <p className={editorClassNames.rejectionTitle}>
        This document was not opened
      </p>
      <p>{IMPORT_REJECTION_REASON[error.code]}</p>
    </div>
  );
}

export function DocxEditor({
  document: source,
  renderImportError,
  ref,
  mode = { kind: "edit" },
  showPageGuides = true,
  zoom,
  defaultZoom = "fit-width",
  onZoomChange,
  fontFallbacks,
  presets,
  commentAuthor = { name: "Anonymous" },
  plugins,
  className,
  style,
  onReady,
  onChange,
}: DocxEditorProps) {
  const readOnly = mode.kind === "readOnly";
  const showToolbar = mode.kind === "edit" && (mode.toolbar ?? true);
  const allowLocking = mode.kind === "edit" && (mode.locking ?? false);
  const bytes = useDocumentBytes(source);
  const opened = useMemo(
    () => (bytes === null ? null : openDocument(bytes)),
    [bytes]
  );
  // Held from the first render on, so a value rebuilt on every render does not rebuild the editor
  const mountedPlugins = useRef(plugins).current;
  // A read-only editor keeps the menu plugins, which stand down while the view is not editable,
  // so that turning editing back on brings the menus back with it
  const mountedContextMenus = useRef(
    mode.kind !== "edit" || (mode.contextMenus ?? true)
  ).current;
  const mountedFontFallbacks = useRef(fontFallbacks).current;
  const layerRef = useRef<HTMLDivElement | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const mountRef = useRef<HTMLDivElement | null>(null);
  const keptState = useRef<{ of: OpenedDocument; state: EditorState } | null>(
    null
  );
  const viewRef = useRef<EditorView | null>(null);
  const [live, setLive] = useState<LiveEditor | null>(null);
  const [commentComposerOpen, setCommentComposerOpen] = useState(false);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [uncontrolledZoom, setUncontrolledZoom] = useState<DocxEditorZoom>(() =>
    normalizeZoom(defaultZoom)
  );
  const latestOnReady = useLatest(onReady);
  const latestOnChange = useLatest(onChange);
  const selectedZoom = normalizeZoom(zoom ?? uncontrolledZoom);
  const page =
    opened?.status === "opened"
      ? pagePixels(opened.session.geometry)
      : A4_PAGE_PIXELS;
  const fitWidth = useFitWidthZoom(rootRef, page.pageWidth);
  const effectiveZoom = selectedZoom === "fit-width" ? fitWidth : selectedZoom;
  const changeZoom = (next: DocxEditorZoom) => {
    const normalized = normalizeZoom(next);
    if (zoom === undefined) setUncontrolledZoom(normalized);
    onZoomChange?.(normalized);
  };

  // Mounted right after render, so a consumer can read the ref straight from its own effect
  useLayoutEffect(() => {
    const mount = mountRef.current;
    if (!mount || opened?.status !== "opened") return;

    const kept = keptState.current;
    const view = createEditorView({
      mount,
      state:
        kept?.of === opened
          ? kept.state
          : createEditorState(opened.doc, {
              numbering: opened.numbering,
              styles: opened.session.styles,
              defaults: opened.session.defaults,
              paragraphDefaults: opened.session.paragraphDefaults,
              canStartNewList: opened.session.numberingPartPath !== null,
              consumerPlugins: mountedPlugins,
              paragraphStyles: opened.session.paragraphStyles,
              contextMenus: mountedContextMenus,
              geometry: opened.session.geometry,
              defaultTabStopPt: opened.session.defaultTabStopPt,
              reservedCommentIds: opened.session.comments.byId.keys(),
              reservedCommentParaIds: [
                ...opened.session.comments.ordered.flatMap((comment) =>
                  comment.paraId === null ? [] : [comment.paraId]
                ),
                ...opened.session.comments.extendedOrdered.map(
                  (extension) => extension.paraId
                ),
              ],
            }),
      defaults: opened.session.defaults,
      geometry: opened.session.geometry,
      readOnly,
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
    readOnly,
    mountedContextMenus,
    mountedFontFallbacks,
    mountedPlugins,
    latestOnReady,
    latestOnChange,
  ]);

  // biome-ignore lint/correctness/useExhaustiveDependencies: when readOnly changes the effect above builds a new view, and the handle has to be rebuilt around that new view
  useImperativeHandle<
    DocxEditorHandle | null,
    DocxEditorHandle | null
  >(ref, () => {
    const view = viewRef.current;
    if (!view || opened?.status !== "opened") return null;
    const session = opened.session;
    return { view, exportBytes: () => exportDocx(view.state.doc, session) };
  }, [opened, readOnly]);

  const overlay = usePageLayout({
    view: live?.view ?? null,
    layer: layerRef,
    enabled: showPageGuides,
    revision: live?.state.doc,
    geometry: opened?.status === "opened" ? opened.session.geometry : undefined,
  });

  if (opened?.status === "rejected") {
    return renderImportError ? (
      renderImportError(opened.error)
    ) : (
      <ImportRejection error={opened.error} />
    );
  }

  // Only ever one menu at a time. Where both plugins hold a point, the text menu is the one the
  // last right click opened, so it is the one drawn.
  // Where the menu plugins were left out, no point is ever held and neither menu is drawn.
  const textAnchor = live && !readOnly ? textMenuAnchor(live.state) : null;
  const tableAnchor =
    live && !readOnly && !textAnchor ? tableMenuAnchor(live.state) : null;

  // The card that says where a link points stands down while the panel that changes one is open, and
  // while an IME is composing: a box appearing and moving under a composition is the churn the page
  // measurement documents avoiding (`page/usePageLayout`)
  const linkAtCursor =
    live && !live.view.composing && !isLinkPanelOpen(live.state)
      ? activeLinkSpan(live.state)
      : null;
  const comments =
    live?.state === undefined ? [] : documentComments(live.state);
  const hasUnresolvedComments = comments.some((comment) => !comment.resolved);
  const effectiveComposerOpen = commentComposerOpen && !readOnly;
  const showComments =
    live !== null &&
    (commentsOpen || effectiveComposerOpen || hasUnresolvedComments);
  const commentsPanel = live && showComments && (
    <CommentsPanel
      view={live.view}
      state={live.state}
      readOnly={readOnly}
      composerOpen={effectiveComposerOpen}
      author={commentAuthor}
      closeComposer={() => setCommentComposerOpen(false)}
      scrollContainer={rootRef.current}
      allCommentsOpen={commentsOpen}
    />
  );

  return (
    <div
      className={[editorClassNames.frame, className].filter(Boolean).join(" ")}
      style={style}
    >
      {live && showToolbar && (
        <Toolbar
          view={live.view}
          state={live.state}
          fontFallbacks={mountedFontFallbacks}
          presets={presets}
          commentsOpen={commentsOpen}
          onToggleComments={() => {
            setCommentsOpen((open) => {
              if (open) setCommentComposerOpen(false);
              return !open;
            });
          }}
          zoom={selectedZoom}
          onZoomChange={changeZoom}
        />
      )}
      <div
        className={editorClassNames.workspace}
        data-comments={showComments ? "visible" : undefined}
      >
        {live && readOnly && comments.length > 0 && (
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
          {/* The paper and the page marks overlaid on it share one positioning origin */}
          <div
            ref={layerRef}
            className={editorClassNames.pageLayer}
            style={{ zoom: effectiveZoom }}
          >
            <div ref={mountRef} />
            {overlay && (
              <PageGuides
                overlay={overlay}
                headersFooters={
                  opened?.status === "opened"
                    ? opened.session.headersFooters
                    : undefined
                }
              />
            )}
            {live && opened?.status === "opened" && (
              <NotesPanel state={live.state} pageWidth={page.pageWidth} />
            )}
          </div>
          {!commentsOpen && commentsPanel}
        </div>
        {commentsOpen && commentsPanel}
      </div>
      {live && !readOnly && isLinkPanelOpen(live.state) && (
        <LinkPanel view={live.view} state={live.state} />
      )}
      {live && linkAtCursor && (
        <LinkCard
          // Escape hides the card over the link it was pressed on; another link is another card
          key={linkAtCursor.from}
          view={live.view}
          state={live.state}
          link={linkAtCursor}
          readOnly={readOnly}
        />
      )}
      {live && textAnchor && (
        <TextMenu
          view={live.view}
          state={live.state}
          anchor={textAnchor}
          allowLocking={allowLocking}
          onAddComment={() => {
            setCommentComposerOpen(true);
          }}
        />
      )}
      {live && tableAnchor && (
        <TableMenu
          view={live.view}
          state={live.state}
          anchor={tableAnchor}
          allowLocking={allowLocking}
        />
      )}
    </div>
  );
}
