import type { EditorView } from "prosemirror-view";
import { readHtmlSlice, withPastedContent } from "./clipboard/htmlReader";
import { readContextOf } from "./clipboard/readContext";
import { insertPlainText } from "./plainText";

export function insertRichHtml(view: EditorView, source: string): boolean {
  const content = readHtmlSlice(
    readContextOf(view.state, view.dom.ownerDocument),
    source
  );
  if (content === null) return false;
  view.dispatch(
    withPastedContent(
      view.state.tr.replaceSelection(content.slice),
      content
    ).scrollIntoView()
  );
  return true;
}

export function insertClipboardData(
  view: EditorView,
  data: { html?: string; text?: string }
): void {
  if (insertRichHtml(view, data.html ?? "")) return;
  insertPlainText(view, data.text ?? "");
}
