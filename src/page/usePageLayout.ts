/**
 * Measures the sheet, moves straddling blocks down to the next page, and hands the
 * page-boundary positions over to the view. It remeasures whenever the text changes or the
 * box is resized.
 *
 * A measurement hands new pushes and break spaces to the view (`./pageDecorations`), which is a
 * decoration transaction against the very paragraph an IME may be composing in, and a composition
 * in Japanese or Chinese stays open across a whole clause while the text grows line by line. So no
 * measurement is taken while a composition is open: the frame is taken again until it is over.
 */

import type { EditorView } from "prosemirror-view";
import {
  type RefObject,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import type { PageGeometry } from "../docx/pageGeometry";
import { editorCssVariables } from "../styles/classNames";
import { measureSheet } from "./measureBlocks";
import {
  setPageBreakSpaces,
  setPagePushes,
  setTableContinuations,
} from "./pageDecorations";
import {
  A4_PAGE_PIXELS,
  PAGE_SPLIT_PX,
  pageLayout,
  pagePixels,
} from "./pageLayout";

/** One place where a page parts from the next. The position is measured on the sheet */
export interface PageMark {
  page: number;
  top: number;
  /** The height of the gap between pages. 0 where the text crosses the boundary */
  height: number;
  crossed: boolean;
}

/** The number laid on the corner of a page */
export interface PageBadge {
  page: number;
  top: number;
  exactPage: boolean;
}

/** The paper area of one visual page, used to place its header and footer stories. */
export interface PageFace {
  page: number;
  headerTop: number;
  footerTop: number;
  left: number;
  width: number;
  crossed: boolean;
}

export interface PageOverlay {
  left: number;
  top: number;
  width: number;
  /** The sheet height with the last page filled out in full */
  sheetHeight: number;
  marks: PageMark[];
  badges: PageBadge[];
  pages: PageFace[];
}

/** How far inside the page's top corner the number is seated */
const BADGE_INSET_PX = 8;

interface PageLayoutOptions {
  view: EditorView | null;
  layer: RefObject<HTMLElement | null>;
  enabled: boolean;
  /** A value that differs every time the text changes */
  revision: unknown;
  /** The paper the open document names. A4 where a document names none */
  geometry?: PageGeometry;
}

/**
 * However many times it is called, the calls are coalesced into a single next frame.
 *
 * While `hold` answers true the frame is taken again instead, so calls made across a stretch that
 * has to be waited out still end in one single run once the wait is over.
 */
function useFrameThrottle(run: () => void, hold: () => boolean): () => void {
  const latest = useRef({ run, hold });
  const frame = useRef(0);

  useEffect(() => {
    latest.current = { run, hold };
  });

  useEffect(() => {
    return () => {
      if (frame.current !== 0) {
        cancelAnimationFrame(frame.current);
        // StrictMode reuses this hook instance across its simulated remount, so the handle has to
        // be cleared as well: a cancelled one left behind reads as a frame already taken and no
        // further call is ever scheduled
        frame.current = 0;
      }
    };
  }, []);

  return useCallback(() => {
    if (frame.current !== 0) return;
    const take = () => {
      frame.current = requestAnimationFrame(() => {
        if (latest.current.hold()) {
          take();
          return;
        }
        frame.current = 0;
        latest.current.run();
      });
    };
    take();
  }, []);
}

/** When the positions and the count are the same, there is nothing to redraw */
function sameOverlay(a: PageOverlay | null, b: PageOverlay): boolean {
  return (
    a !== null &&
    a.left === b.left &&
    a.top === b.top &&
    a.width === b.width &&
    a.sheetHeight === b.sheetHeight &&
    JSON.stringify(a.marks) === JSON.stringify(b.marks) &&
    JSON.stringify(a.badges) === JSON.stringify(b.badges) &&
    JSON.stringify(a.pages) === JSON.stringify(b.pages)
  );
}

export function usePageLayout({
  view,
  layer,
  enabled,
  revision,
  geometry,
}: PageLayoutOptions): PageOverlay | null {
  const [overlay, setOverlay] = useState<PageOverlay | null>(null);
  // The paper is fixed the moment the document is opened, so the pixels are worked out once
  const page = useMemo(
    () => (geometry ? pagePixels(geometry) : A4_PAGE_PIXELS),
    [geometry]
  );

  const remeasure = useFrameThrottle(
    () => {
      const box = layer.current;
      if (!view || !box || !enabled) return;

      const measured = measureSheet(view, box);
      const layout = pageLayout({
        blocks: measured.blocks,
        pageBodyHeight: page.bodyHeight,
        pageStep: page.pageStep,
      });
      setPagePushes(view, layout.pushes);
      setPageBreakSpaces(view, layout.spaces);
      setTableContinuations(view, layout.tableContinuations);

      // Stretch the sheet to the number of pages so the last one also looks like a full page
      const sheetHeight =
        measured.contentTop + layout.bodyHeight + measured.contentBottom;
      box.style.setProperty(editorCssVariables.sheetHeight, `${sheetHeight}px`);

      const next: PageOverlay = {
        left: measured.left,
        top: measured.top,
        width: measured.width,
        sheetHeight,
        marks: layout.splits.map((split) => ({
          page: split.page,
          top:
            measured.contentTop +
            split.y +
            (split.crossed ? 0 : page.marginBottom),
          height: split.crossed ? 0 : PAGE_SPLIT_PX,
          crossed: split.crossed,
        })),
        // The number is laid on the page's top corner. A page crossed into has no margin,
        // so the place it was split at is its corner
        badges: layout.pages.map((start) => ({
          page: start.page,
          top:
            measured.contentTop +
            start.bodyStart -
            (start.crossed ? 0 : page.marginTop) +
            BADGE_INSET_PX,
          exactPage: start.exactPage,
        })),
        pages: layout.pages.map((start) => {
          const paperTop =
            measured.contentTop +
            start.bodyStart -
            (start.crossed ? 0 : page.marginTop);
          return {
            page: start.page,
            headerTop: paperTop + page.marginTop / 2,
            footerTop: paperTop + page.pageHeight - page.marginBottom / 2,
            left: page.marginLeft,
            width: page.bodyWidth,
            crossed: start.crossed,
          };
        }),
      };
      setOverlay((previous) => (sameOverlay(previous, next) ? previous : next));
    },
    () => view?.composing === true
  );

  // When the text changes, the block heights change with it
  // biome-ignore lint/correctness/useExhaustiveDependencies: revision is the remeasure trigger, compared by identity and never read - the doc must be read from view.state at measure time, not from this closure
  useEffect(() => {
    if (enabled) remeasure();
  }, [remeasure, revision, enabled]);

  useEffect(() => {
    if (enabled) return;
    layer.current?.style.removeProperty(editorCssVariables.sheetHeight);
    if (!view) return;
    setPagePushes(view, []);
    setPageBreakSpaces(view, []);
    setTableContinuations(view, []);
  }, [enabled, layer, view]);

  useEffect(() => {
    const box = layer.current;
    // Where there is no layout (in tests) nothing is ever resized either
    if (!view || !box || !enabled || typeof ResizeObserver === "undefined") {
      return;
    }
    const observer = new ResizeObserver(remeasure);
    observer.observe(view.dom);
    observer.observe(box);
    return () => observer.disconnect();
  }, [view, layer, enabled, remeasure]);

  return enabled ? overlay : null;
}
