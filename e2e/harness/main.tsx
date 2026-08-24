/**
 * The page the composition tests drive.
 *
 * It mounts one editor over the bytes of a fixture named in the query string and hangs a reading
 * surface off `window` (`./api`). A real IME is driven into it over CDP, so everything the tests
 * assert has to be readable from this one page: the document model, the sheet as drawn, and the
 * composition events the browser fired.
 */

import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { DocxEditor } from "../../src/DocxEditor";
import type { DocxBytes } from "../../src/docx/importDocx";
import { lockSelection } from "../../src/editor/commands/lockCommands";
import { lockedMarkOf } from "../../src/schema/locks";
import { editorAttributes } from "../../src/styles/classNames";
import "../../src/styles/editor.css";
import type {
  BlockReport,
  CaretBox,
  CompositionCounts,
  DocxHarness,
} from "./api";
import { longTableFixture } from "./longTableFixture";
import { tabFixture } from "./tabFixture";

const fixtureUrls = import.meta.glob<string>("../../__fixtures__/*.docx", {
  query: "?url",
  import: "default",
  eager: true,
});

const DEFAULT_FIXTURE = "kitchen-sink";

function fixtureUrl(name: string): string {
  const url = fixtureUrls[`../../__fixtures__/${name}.docx`];
  if (url === undefined) {
    const known = Object.keys(fixtureUrls).join(", ");
    throw new Error(`no such fixture: ${name}. known: ${known}`);
  }
  return url;
}

async function loadFixture(name: string): Promise<DocxBytes> {
  const generated = name === "long-table" || name === "tabs";
  const fixture = generated ? DEFAULT_FIXTURE : name;
  const response = await fetch(fixtureUrl(fixture));
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (name === "long-table") return longTableFixture(bytes);
  if (name === "tabs") return tabFixture(bytes);
  return bytes;
}

function askedFixture(): string {
  const asked = new URLSearchParams(window.location.search).get("fixture");
  return asked ?? DEFAULT_FIXTURE;
}

function blockStart(view: EditorView, blockIndex: number): number {
  let at = 0;
  view.state.doc.forEach((_node, pos, index) => {
    if (index === blockIndex) at = pos;
  });
  return at;
}

function blocks(view: EditorView): BlockReport[] {
  const found: BlockReport[] = [];
  view.state.doc.forEach((node, pos, index) => {
    const dom = view.nodeDOM(pos);
    found.push({
      index,
      pos,
      type: node.type.name,
      docText: node.textBetween(0, node.content.size, "\n", "\n"),
      domText: dom instanceof HTMLElement ? (dom.textContent ?? "") : "",
      marker:
        dom instanceof HTMLElement
          ? dom.getAttribute(editorAttributes.listMarker)
          : null,
    });
  });
  return found;
}

function blockHeight(view: EditorView, blockIndex: number): number {
  const dom = view.nodeDOM(blockStart(view, blockIndex));
  return dom instanceof HTMLElement ? dom.getBoundingClientRect().height : 0;
}

/** The pushed blocks as one string, so a test can watch the decorations stand still or move */
function pushes(view: EditorView): string {
  const pushed = view.dom.querySelectorAll(`[${editorAttributes.pagePush}]`);
  return Array.from(pushed, (element) => {
    const gap =
      element instanceof HTMLElement ? element.style.marginBlockStart : "";
    return `${element.getAttribute(editorAttributes.pagePush)}/${gap}`;
  }).join(",");
}

function pagination(view: EditorView): string {
  const tableSpaces = view.dom.querySelectorAll(
    `[${editorAttributes.tablePageSpace}]`
  );
  return [
    pushes(view),
    Array.from(tableSpaces, (element) => {
      const height = element instanceof HTMLElement ? element.style.height : "";
      return `${element.getAttribute(editorAttributes.tablePageSpace)}/${height}`;
    }).join(","),
  ]
    .filter(Boolean)
    .join(",");
}

/** The spaces opened at the page breaks, so a test can watch them stand still or move */
function spaces(view: EditorView): string {
  const opened = view.dom.querySelectorAll(
    `[${editorAttributes.pageBreakSpace}]`
  );
  return Array.from(opened, (element) => {
    const height = element instanceof HTMLElement ? element.style.height : "";
    return `${element.getAttribute(editorAttributes.pageBreakSpace)}/${height}`;
  }).join(",");
}

function caretBox(view: EditorView): CaretBox {
  const at = view.coordsAtPos(view.state.selection.head);
  return { top: at.top, bottom: at.bottom, left: at.left };
}

/** The position of the first table cell in the document, and the table it belongs to */
function firstCell(view: EditorView): { table: number; cell: number } | null {
  let table: number | null = null;
  let cell: number | null = null;
  view.state.doc.descendants((node, pos) => {
    if (cell !== null) return false;
    if (node.type.spec.tableRole === "table" && table === null) table = pos;
    if (node.type.spec.tableRole === "cell") cell = pos;
    return true;
  });
  return table === null || cell === null ? null : { table, cell };
}

function caretInCell(view: EditorView): string {
  const found = firstCell(view);
  if (!found) throw new Error("the fixture holds no table");
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.near(view.state.doc.resolve(found.cell + 1))
    )
  );
  view.focus();
  return view.state.doc.nodeAt(found.cell)?.textContent ?? "";
}

function rightClick(view: EditorView): void {
  const at = view.coordsAtPos(view.state.selection.head);
  const node = view.domAtPos(view.state.selection.head).node;
  const element = node instanceof Element ? node : node.parentElement;
  element?.dispatchEvent(
    new MouseEvent("contextmenu", {
      bubbles: true,
      cancelable: true,
      clientX: at.left,
      clientY: at.top,
    })
  );
}

function tableRows(view: EditorView): number {
  const found = firstCell(view);
  return found === null
    ? 0
    : (view.state.doc.nodeAt(found.table)?.childCount ?? 0);
}

function lockedText(view: EditorView): string {
  let text = "";
  view.state.doc.descendants((node) => {
    if (node.isText && lockedMarkOf(node)) text += node.text ?? "";
    return true;
  });
  return text;
}

function install(view: EditorView): void {
  const counts: CompositionCounts = { start: 0, update: 0, end: 0 };
  view.dom.addEventListener("compositionstart", () => {
    counts.start += 1;
  });
  view.dom.addEventListener("compositionupdate", () => {
    counts.update += 1;
  });
  view.dom.addEventListener("compositionend", () => {
    counts.end += 1;
  });

  const harness: DocxHarness = {
    blocks: () => blocks(view),
    text: () =>
      view.state.doc.textBetween(0, view.state.doc.content.size, "\n", "\n"),
    composing: () => view.composing,
    compositions: () => ({ ...counts }),
    pushes: () => pushes(view),
    pagination: () => pagination(view),
    spaces: () => spaces(view),
    caretBox: () => caretBox(view),
    selection: () => ({
      from: view.state.selection.from,
      to: view.state.selection.to,
      anchor: view.state.selection.anchor,
      head: view.state.selection.head,
    }),
    blockHeight: (blockIndex) => blockHeight(view, blockIndex),
    caretAt: (blockIndex, offset) => {
      const at = blockStart(view, blockIndex) + 1 + offset;
      view.dispatch(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, at))
      );
      view.focus();
      return at;
    },
    selectText: (blockIndex, offset, length) => {
      const from = blockStart(view, blockIndex) + 1 + offset;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, from, from + length)
        )
      );
      view.focus();
    },
    lock: (blockIndex, offset, length) => {
      const from = blockStart(view, blockIndex) + 1 + offset;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, from, from + length)
        )
      );
      return lockSelection(view.state, (tr) => view.dispatch(tr));
    },
    caretInCell: () => caretInCell(view),
    rightClick: () => rightClick(view),
    tableRows: () => tableRows(view),
    lockedText: () => lockedText(view),
  };

  window.docxHarness = harness;
}

function Harness() {
  const [bytes, setBytes] = useState<DocxBytes | null>(null);

  useEffect(() => {
    let live = true;
    loadFixture(askedFixture()).then((loaded) => {
      if (live) setBytes(loaded);
    });
    return () => {
      live = false;
    };
  }, []);

  if (!bytes) return null;
  return (
    <div data-testid="editor">
      <DocxEditor
        document={bytes}
        renderImportError={(error) => (
          <p data-testid="import-error">{error.code}</p>
        )}
        mode={{ kind: "edit", locking: true }}
        onReady={install}
      />
    </div>
  );
}

const root = document.getElementById("root");
if (root) createRoot(root).render(<Harness />);
