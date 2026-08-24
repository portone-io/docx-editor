import type { EditorView } from "prosemirror-view";
import { type RefObject, useLayoutEffect, useRef, useState } from "react";
import type { DocumentComment } from "../../editor/commands/commentCommands";
import { editorClassNames } from "../../styles/classNames";

export const COMPOSER_POSITION = "composer";
const VERTICAL_SCROLLBAR_CLEARANCE = 16;

function estimatedCardHeight(comment: DocumentComment): number {
  return 150 + comment.replies.length * 88;
}

function topAt(view: EditorView, pos: number): number {
  try {
    return view.coordsAtPos(pos).top;
  } catch {
    return 0;
  }
}

function anchorPositions(
  comments: readonly DocumentComment[],
  view: EditorView,
  scrollContainer: HTMLElement,
  rail: HTMLElement,
  composerOpen: boolean,
  selectionPos: number
): { positions: ReadonlyMap<string, number>; bottom: number } {
  const root = scrollContainer.getBoundingClientRect();
  const heights = new Map<string, number>();
  for (const element of rail.querySelectorAll<HTMLElement>(
    "[data-comment-position]"
  )) {
    const key = element.dataset.commentPosition;
    if (key) heights.set(key, element.getBoundingClientRect().height);
  }
  const items = comments.map((comment) => ({
    key: comment.id,
    anchor: topAt(view, comment.from) - root.top + scrollContainer.scrollTop,
    fallbackHeight: estimatedCardHeight(comment),
  }));
  if (composerOpen) {
    items.push({
      key: COMPOSER_POSITION,
      anchor: topAt(view, selectionPos) - root.top + scrollContainer.scrollTop,
      fallbackHeight: 150,
    });
  }
  items.sort((left, right) => left.anchor - right.anchor);
  const positions = new Map<string, number>();
  let next = 64;
  for (const item of items) {
    const top = Math.max(item.anchor, next);
    positions.set(item.key, top);
    next = top + (heights.get(item.key) || item.fallbackHeight) + 12;
  }
  return { positions, bottom: next };
}

interface CommentRailLayoutOptions {
  panel: RefObject<HTMLElement | null>;
  view: EditorView;
  scrollContainer: HTMLElement | null;
  visible: readonly DocumentComment[];
  composerOpen: boolean;
  composerSelectionPos: number;
  allCommentsOpen: boolean;
}

export function useCommentRailLayout({
  panel,
  view,
  scrollContainer,
  visible,
  composerOpen,
  composerSelectionPos,
  allCommentsOpen,
}: CommentRailLayoutOptions): {
  positions: ReadonlyMap<string, number>;
  canvasHeight: number;
} {
  const wasAllCommentsOpen = useRef(false);
  const [positions, setPositions] = useState<ReadonlyMap<string, number>>(
    new Map()
  );
  const [canvasHeight, setCanvasHeight] = useState(0);

  useLayoutEffect(() => {
    if (allCommentsOpen && !wasAllCommentsOpen.current && panel.current) {
      panel.current.scrollTop = 0;
    }
    wasAllCommentsOpen.current = allCommentsOpen;
  }, [allCommentsOpen, panel]);

  useLayoutEffect(() => {
    const rail = panel.current;
    if (!rail || !scrollContainer) return;
    const frameWindow = scrollContainer.ownerDocument.defaultView;
    const pageLayer = view.dom.closest<HTMLElement>(
      `.${editorClassNames.pageLayer}`
    );
    if (allCommentsOpen) {
      pageLayer?.style.removeProperty("padding-right");
      rail.style.removeProperty("left");
      setPositions(new Map());
      setCanvasHeight(0);
      const wheelAllComments = (event: WheelEvent) => {
        event.preventDefault();
        const before = rail.scrollTop;
        const max = Math.max(0, rail.scrollHeight - rail.clientHeight);
        rail.scrollTop = Math.min(max, Math.max(0, before + event.deltaY));
        scrollContainer.scrollTop += event.deltaY - (rail.scrollTop - before);
        scrollContainer.scrollLeft += event.deltaX;
      };
      rail.addEventListener("wheel", wheelAllComments, { passive: false });
      return () => rail.removeEventListener("wheel", wheelAllComments);
    }
    const reserveCommentWidth = () => {
      if (!pageLayer) return;
      const drawnWidth = rail.getBoundingClientRect().width || rail.offsetWidth;
      const computedZoom = Number.parseFloat(
        frameWindow?.getComputedStyle(pageLayer).zoom ?? ""
      );
      const scale =
        Number.isFinite(computedZoom) && computedZoom > 0 ? computedZoom : 1;
      pageLayer.style.paddingRight = `${(drawnWidth + 16 + VERTICAL_SCROLLBAR_CLEARANCE) / scale}px`;
    };
    const placeBesidePaper = () => {
      reserveCommentWidth();
      const viewport = scrollContainer.getBoundingClientRect();
      const paper = view.dom.getBoundingClientRect();
      const desired =
        paper.right - viewport.left + scrollContainer.scrollLeft + 16;
      rail.style.left = `${Math.max(16, desired)}px`;
    };
    const update = () => {
      const layout = anchorPositions(
        visible,
        view,
        scrollContainer,
        rail,
        composerOpen,
        composerSelectionPos
      );
      setPositions(layout.positions);
      const bottom =
        visible.length > 0 || composerOpen ? layout.bottom + 12 : 0;
      setCanvasHeight(Math.max(scrollContainer.scrollHeight, bottom));
    };
    const wheelDocument = (event: WheelEvent) => {
      event.preventDefault();
      scrollContainer.scrollTop += event.deltaY;
      scrollContainer.scrollLeft += event.deltaX;
    };
    const placeAndUpdate = () => {
      placeBesidePaper();
      update();
    };
    placeAndUpdate();
    let settledFrame: number | undefined;
    const frame = frameWindow?.requestAnimationFrame(() => {
      placeAndUpdate();
      settledFrame = frameWindow.requestAnimationFrame(() => {
        placeAndUpdate();
      });
    });
    frameWindow?.addEventListener("resize", placeAndUpdate, {
      passive: true,
    });
    rail.addEventListener("wheel", wheelDocument, { passive: false });
    const resize =
      typeof ResizeObserver === "undefined"
        ? null
        : new ResizeObserver(placeAndUpdate);
    resize?.observe(scrollContainer);
    resize?.observe(view.dom);
    resize?.observe(rail);
    for (const element of rail.querySelectorAll<HTMLElement>(
      "[data-comment-position]"
    )) {
      resize?.observe(element);
    }
    const mutations =
      typeof MutationObserver === "undefined"
        ? null
        : new MutationObserver(() => {
            placeAndUpdate();
            for (const element of rail.querySelectorAll<HTMLElement>(
              "[data-comment-position]"
            )) {
              resize?.observe(element);
            }
          });
    mutations?.observe(rail, { childList: true, subtree: true });
    return () => {
      if (frame !== undefined) frameWindow?.cancelAnimationFrame(frame);
      if (settledFrame !== undefined) {
        frameWindow?.cancelAnimationFrame(settledFrame);
      }
      frameWindow?.removeEventListener("resize", placeAndUpdate);
      rail.removeEventListener("wheel", wheelDocument);
      mutations?.disconnect();
      resize?.disconnect();
      pageLayer?.style.removeProperty("padding-right");
    };
  }, [
    allCommentsOpen,
    composerOpen,
    composerSelectionPos,
    panel,
    scrollContainer,
    view,
    visible,
  ]);

  return { positions, canvasHeight };
}
