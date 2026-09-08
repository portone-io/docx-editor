/**
 * Draws the gaps between pages on the sheet.
 *
 * Straddling blocks have already been moved to the next page, so there is no text under
 * these bands.
 * Only the places crossed by a block taller than one page are marked over the text, with
 * a faint dotted line.
 * The box only paints, so the mouse still reaches the text beneath it.
 *
 * Each page draws the header and footer of the section it opens with, since a document names its
 * headers section by section (`docx/headersFooters`).
 */

import type { ReactElement } from "react";
import {
  displayPageNumber,
  type HeadersFooters,
  headerFooterOn,
  headerFooterText,
} from "../docx/headersFooters";
import { editorClassNames } from "../styles/classNames";
import type { PageFace, PageOverlay } from "./usePageLayout";

export function PageGuides({
  overlay,
  headersFootersFor,
}: {
  overlay: PageOverlay;
  /** The stories the section this page belongs to shows. Left out while no document is open */
  headersFootersFor?: (page: PageFace) => HeadersFooters | null;
}): ReactElement {
  const totalPages = overlay.pages.length;
  return (
    <div
      className={editorClassNames.pageGuides}
      style={{
        left: `${overlay.left}px`,
        top: `${overlay.top}px`,
        width: `${overlay.width}px`,
        height: `${overlay.sheetHeight}px`,
      }}
      aria-hidden="true"
    >
      {overlay.marks.map((mark) => (
        <div
          key={mark.page}
          className={
            mark.crossed
              ? editorClassNames.pageCrossed
              : editorClassNames.pageSplit
          }
          style={{
            top: `${mark.top}px`,
            height: mark.height > 0 ? `${mark.height}px` : undefined,
          }}
        />
      ))}
      {headersFootersFor &&
        overlay.pages.flatMap((page) => {
          const headersFooters = page.crossed ? null : headersFootersFor(page);
          if (headersFooters === null) return [];
          const number = displayPageNumber(headersFooters, page.page);
          const header = headerFooterOn(
            headersFooters.headers,
            headersFooters,
            page.pageInSection
          );
          const footer = headerFooterOn(
            headersFooters.footers,
            headersFooters,
            page.pageInSection
          );
          return [
            header === null ? null : (
              <div
                key={`header-${page.page}`}
                className={editorClassNames.pageHeader}
                style={{
                  top: `${page.headerTop}px`,
                  left: `${page.left}px`,
                  width: `${page.width}px`,
                  textAlign: header.align ?? undefined,
                }}
                title={`Page ${number} header`}
              >
                {headerFooterText(header.story, number, totalPages)}
              </div>
            ),
            footer === null ? null : (
              <div
                key={`footer-${page.page}`}
                className={editorClassNames.pageFooter}
                style={{
                  top: `${page.footerTop}px`,
                  left: `${page.left}px`,
                  width: `${page.width}px`,
                  textAlign: footer.align ?? undefined,
                }}
                title={`Page ${number} footer`}
              >
                {headerFooterText(footer.story, number, totalPages)}
              </div>
            ),
          ];
        })}
    </div>
  );
}
