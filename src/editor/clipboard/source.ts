/**
 * Which application wrote the markup on the clipboard.
 *
 * Every writer of HTML says so somewhere - in a generator meta element, in a class prefix, in the
 * id it wraps a copy in - and what it says decides which readings apply: Word's list paragraphs
 * are lists rather than paragraphs beginning with a bullet character, and its `<o:p>` markers are
 * nothing at all. A reading recognizing no writer stays with what the markup itself states.
 */

import { editorClassNames } from "../../styles/classNames";
import { INTERNAL_TOKEN_ATTRIBUTE } from "./internalChannel";

/** The writers whose markup is read differently from plain HTML */
export type HtmlSource =
  | "editor"
  | "word"
  | "google-docs"
  | "libreoffice"
  | "unknown";

/**
 * What each writer is known by, tried in order.
 *
 * A browser keeps a copy's own markup whole, so these are read off the copy rather than off a
 * clipboard header nobody but the copying application sees. `prosemirror-view` drops the meta
 * elements standing ahead of the markup and keeps the ones a copy carries inside its `<head>`,
 * which is where Word and LibreOffice write theirs.
 */
const SIGNATURES: readonly (readonly [HtmlSource, string])[] = [
  ["editor", `[${INTERNAL_TOKEN_ATTRIBUTE}], .${editorClassNames.paragraph}`],
  ["word", 'meta[name="Generator" i][content^="Microsoft Word" i], .MsoNormal'],
  ["google-docs", '[id^="docs-internal-guid-"]'],
  [
    "libreoffice",
    'meta[name="generator" i][content*="LibreOffice" i], meta[name="generator" i][content*="OpenOffice" i]',
  ],
];

/** The application the markup under this element was written by */
export function detectHtmlSource(root: ParentNode): HtmlSource {
  for (const [source, selector] of SIGNATURES) {
    if (root.querySelector(selector) !== null) return source;
  }
  return "unknown";
}
