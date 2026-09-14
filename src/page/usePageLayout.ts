/**
 * Measures the sheet, moves straddling blocks down to the next page, and hands the
 * page-boundary positions over to the view. It remeasures whenever the text changes or the
 * box is resized.
 *
 * A measurement hands its marks to the view (`./pageDecorations`) in one decoration transaction
 * against the very paragraph an IME may be composing in, and a composition
 * in Japanese or Chinese stays open across a whole clause while the text grows line by line. So no
 * measurement is taken while a composition is open: the frame is taken again until it is over.
 */

import type { EditorView } from "prosemirror-view";
import {
  type RefObject,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import { editorCssVariables } from "../styles/classNames";
import type { DemandBand } from "./demands";
import { measureSheet } from "./measureBlocks";
import { setPageMarks } from "./pageDecorations";
import {
  A4_SECTION_PIXELS,
  PAGE_SPLIT_PX,
  type PageReservation,
  pageLayout,
  type SectionPixels,
  sectionPaperAt,
} from "./pageLayout";

/** One place where a page parts from the next. The position is measured on the sheet */
export interface PageMark {
  page: number;
  top: number;
  /** The height of the gap between pages. 0 where the text crosses the boundary */
  height: number;
  crossed: boolean;
}

/** The room one page keeps at the foot of its body for one band, measured on the sheet */
export interface ReservedRoom {
  band: string;
  /** What the page keeps there, in the order it took them (`page/pageLayout`) */
  ids: readonly string[];
  top: number;
  /**
   * The height the page gives the band, down to the end of its body. Less than the band asks for
   * where it holds more than the page could give it
   */
  height: number;
  /** Whether the band asks for more than that height */
  clipped: boolean;
}

/** The paper area of one visual page, used to place its header and footer stories. */
export interface PageFace {
  page: number;
  /** The position of the block this page opens with, which says which section it belongs to */
  pos: number;
  /**
   * The place this page takes within that section, counted from 1 again at every section, which is
   * what its header and footer variant is chosen by (`docx/headersFooters`)
   */
  pageInSection: number;
  headerTop: number;
  footerTop: number;
  left: number;
  width: number;
  crossed: boolean;
  /** What the page keeps at the foot of its body, one entry per band, top to bottom */
  reserved: readonly ReservedRoom[];
}

export interface PageOverlay {
  left: number;
  top: number;
  width: number;
  /** The sheet height with the last page filled out in full */
  sheetHeight: number;
  marks: PageMark[];
  pages: PageFace[];
}

interface PageLayoutOptions {
  view: EditorView | null;
  layer: RefObject<HTMLElement | null>;
  enabled: boolean;
  /** A value that differs every time the text changes */
  revision: unknown;
  /**
   * The paper of each section of the open document (`page/pageLayout`), read once per document
   * change by the caller. A4 where a document names none, and where none is open
   */
  sections?: readonly SectionPixels[];
  /**
   * The bands the demands of the text are kept in (`page/demands`), by name. Left out, the pages
   * are the ones the blocks alone come to, which is also what a document asking for no room
   * should be laid out through. A new value lays the pages out again
   */
  bands?: ReadonlyMap<string, DemandBand>;
  /**
   * Whether anything drawn over the sheet besides the body is composing, such as the view over a
   * note being edited.
   *
   * A measurement moves the room a band is kept in, and what is drawn in that room is drawn again
   * wherever it lands, which would take an open composition down with it. So the frame is taken
   * again until that composition is over, the way it is for the body's own.
   */
  composing?: () => boolean;
}

const NO_ROOM: readonly ReservedRoom[] = [];

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
    JSON.stringify(a.pages) === JSON.stringify(b.pages)
  );
}

export function usePageLayout({
  view,
  layer,
  enabled,
  revision,
  sections,
  bands,
  composing,
}: PageLayoutOptions): PageOverlay | null {
  const [overlay, setOverlay] = useState<PageOverlay | null>(null);
  const papers =
    sections === undefined || sections.length === 0
      ? A4_SECTION_PIXELS
      : sections;

  const remeasure = useFrameThrottle(
    () => {
      const box = layer.current;
      if (!view || !box || !enabled) return;

      const measured = measureSheet(view, box);
      const layout = pageLayout({
        blocks: measured.blocks,
        sections: papers,
        bands,
      });
      const roomOn = new Map<number, PageReservation[]>();
      for (const room of layout.reserved) {
        roomOn.set(room.page, [...(roomOn.get(room.page) ?? []), room]);
      }
      /** The paper of the page that block opens, which is the paper of its own section */
      const paperOf = (pos: number) => sectionPaperAt(papers, pos);
      // One sheet is drawn at one width, the first section's (`styles/editor.css`), so where a
      // page stands across it is that paper's while how tall it stands is its own section's
      const sheet = paperOf(0);
      setPageMarks(view, { pushes: layout.pushes, cuts: layout.cuts });

      // Stretch the sheet to the number of pages so the last one also looks like a full page.
      // The padding it is drawn with is the first section's, so where the document ends on a
      // deeper bottom margin the sheet is stretched to that margin instead: otherwise the last
      // page's footer would be drawn past the end of the paper
      const sheetHeight =
        measured.contentTop +
        layout.bodyHeight +
        Math.max(measured.contentBottom, layout.marginBottom);
      box.style.setProperty(editorCssVariables.sheetHeight, `${sheetHeight}px`);

      const next: PageOverlay = {
        left: measured.left,
        top: measured.top,
        width: measured.width,
        sheetHeight,
        // A gap opens below the page that ends at this place, so it is that page's own margin
        // the band is drawn under
        marks: layout.splits.map((split) => ({
          page: split.page,
          top:
            measured.contentTop +
            split.y +
            (split.crossed
              ? 0
              : paperOf(layout.pages[split.page - 2]?.pos ?? 0).marginBottom),
          height: split.crossed ? 0 : PAGE_SPLIT_PX,
          crossed: split.crossed,
        })),
        // A page crossed into has no margin, so the place it was split at is its top corner
        pages: layout.pages.map((start) => {
          const paper = paperOf(start.pos);
          const paperTop =
            measured.contentTop +
            start.bodyStart -
            (start.crossed ? 0 : paper.marginTop);
          const bodyBottom =
            measured.contentTop + start.bodyStart + paper.bodyHeight;
          return {
            page: start.page,
            pos: start.pos,
            pageInSection: start.pageInSection,
            headerTop: paperTop + paper.marginTop / 2,
            footerTop: paperTop + paper.pageHeight - paper.marginBottom / 2,
            left: sheet.marginLeft,
            width: sheet.bodyWidth,
            crossed: start.crossed,
            reserved:
              roomOn.get(start.page)?.map((room) => {
                const top = measured.contentTop + room.top;
                const height = Math.max(0, bodyBottom - top);
                return {
                  band: room.band,
                  ids: room.ids,
                  top,
                  height: Math.min(room.height, height),
                  clipped: room.height > height,
                };
              }) ?? NO_ROOM,
          };
        }),
      };
      setOverlay((previous) => (sameOverlay(previous, next) ? previous : next));
    },
    () => view?.composing === true || composing?.() === true
  );

  // When the text changes, the block heights change with it
  // biome-ignore lint/correctness/useExhaustiveDependencies: revision is the remeasure trigger, compared by identity and never read - the doc must be read from view.state at measure time, not from this closure
  useEffect(() => {
    if (enabled) remeasure();
  }, [remeasure, revision, enabled, bands]);

  // Everything the last measurement left behind goes with the pages, the overlay included: held
  // on, it would be handed out again the moment a consumer turns them back on, and the guides and
  // the footnotes standing over them would be drawn at positions measured before they went off
  useEffect(() => {
    if (enabled) return;
    setOverlay(null);
    layer.current?.style.removeProperty(editorCssVariables.sheetHeight);
    if (!view) return;
    setPageMarks(view, { pushes: [], cuts: [] });
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
