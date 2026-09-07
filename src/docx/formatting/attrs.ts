/**
 * The derived attrs a paragraph and the run marks inside it carry, worked out from the resolved
 * hierarchy. Import, the plugin that formats the paragraphs an edit built, the paragraph writers
 * and the clipboard reader all take them from here, so a paragraph is drawn the same whichever of
 * them last touched it.
 */

import type { Mark } from "prosemirror-model";
import type { ParagraphFormat, RunFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import type { FormattingContext } from "./context";
import {
  type ParagraphPlacement,
  type ResolvedParagraph,
  resolveParagraph,
  resolveRun,
} from "./resolve";

/** The two derived paragraph attrs (`format` and `styleRun` in `schema`) */
export interface ParagraphAttrs {
  format: ParagraphFormat | null;
  styleRun: RunFormat | null;
}

export function paragraphAttrsOf(paragraph: ResolvedParagraph): ParagraphAttrs {
  return { format: paragraph.format, styleRun: paragraph.styleRun };
}

export function paragraphAttrsFor(
  pPr: string | null,
  context: FormattingContext,
  placement: ParagraphPlacement | null = null
): ParagraphAttrs {
  return paragraphAttrsOf(resolveParagraph(pPr, context, placement));
}

/**
 * The run mark an inline wears under the resolved paragraph, or null when it needs none.
 *
 * A mark already there keeps the formatting it wrote down and has its display values read
 * again, since they were laid over whatever the paragraph wore before. Text carrying no mark
 * gets one holding nothing but the values the paragraph passes down, so it is drawn the way
 * import draws a bare `w:r`; such a mark writes no `w:rPr` and changes nothing on the way out.
 * An inline that is not text and carries no mark is left as it is.
 */
export function runMarkUnder(
  mark: Mark | null,
  isText: boolean,
  paragraph: ResolvedParagraph,
  context: FormattingContext
): Mark | null {
  if (mark) {
    const rPr: unknown = mark.attrs.rPr;
    const format = resolveRun(
      typeof rPr === "string" ? rPr : null,
      paragraph,
      context
    );
    return mark.type.create({ ...mark.attrs, format });
  }
  if (!isText) return null;
  const format = resolveRun(null, paragraph, context);
  return format === null
    ? null
    : docxSchema.marks.run.create({ rPr: null, rAttrs: null, format });
}
