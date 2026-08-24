/**
 * Draws the gaps between pages on the sheet, plus a number for each page.
 *
 * Straddling blocks have already been moved to the next page, so there is no text under
 * these bands.
 * Only the places crossed by a block taller than one page are marked over the text, with
 * a faint dotted line.
 * The box only paints, so the mouse still reaches the text beneath it.
 */

import {
  displayPageNumber,
  type HeadersFooters,
  headerFooterAlign,
  headerFooterText,
} from "../docx/headersFooters";
import { editorClassNames } from "../styles/classNames";
import type { PageBadge, PageOverlay } from "./usePageLayout";

/** Only the number itself is shown; that it is an estimate is left to the tooltip */
function badgeTitle(badge: PageBadge): string {
  return badge.exactPage
    ? `Page ${badge.page}`
    : `Page ${badge.page} (approximate)`;
}

export function PageGuides({
  overlay,
  headersFooters,
}: {
  overlay: PageOverlay;
  headersFooters?: HeadersFooters;
}) {
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
      {overlay.badges.map((badge) => (
        <span
          key={badge.page}
          className={editorClassNames.pageBadge}
          style={{ top: `${badge.top}px` }}
          title={badgeTitle(badge)}
        >
          {badge.page}
        </span>
      ))}
      {headersFooters &&
        overlay.pages.flatMap((page) => {
          if (page.crossed) return [];
          const number = displayPageNumber(headersFooters, page.page);
          const header = headerFooterText(
            headersFooters.headers,
            headersFooters,
            page.page,
            totalPages
          );
          const footer = headerFooterText(
            headersFooters.footers,
            headersFooters,
            page.page,
            totalPages
          );
          const headerAlign = headerFooterAlign(
            headersFooters.headers,
            headersFooters,
            page.page
          );
          const footerAlign = headerFooterAlign(
            headersFooters.footers,
            headersFooters,
            page.page
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
                  textAlign: headerAlign ?? undefined,
                }}
                title={`Page ${number} header`}
              >
                {header}
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
                  textAlign: footerAlign ?? undefined,
                }}
                title={`Page ${number} footer`}
              >
                {footer}
              </div>
            ),
          ];
        })}
    </div>
  );
}
