import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { docxSchema } from "../../schema";
import { editorClassNames } from "../../styles/classNames";

type TabSide = "before" | "after";

interface TabBoundary {
  position: number;
  side: TabSide;
}

function isTab(node: PMNode | null, edge: "start" | "end"): boolean {
  if (
    !node?.isText ||
    !node.marks.some((mark) => mark.type === docxSchema.marks.tab)
  ) {
    return false;
  }
  return edge === "start"
    ? node.text?.startsWith("\t") === true
    : node.text?.endsWith("\t") === true;
}

function tabBoundary(view: EditorView): TabBoundary | null {
  const { selection } = view.state;
  if (!selection.empty) return null;
  const position = selection.head;
  const resolved = view.state.doc.resolve(position);
  if (isTab(resolved.nodeAfter, "start")) return { position, side: "before" };
  if (isTab(resolved.nodeBefore, "end")) {
    return { position: position - 1, side: "after" };
  }
  return null;
}

function tabSlot(view: EditorView, position: number): HTMLElement | null {
  return view.dom.querySelector<HTMLElement>(
    `.${editorClassNames.tabSlot}[data-tab-position="${position}"]`
  );
}

/** Draws a caret from the model boundary because browsers disagree on a tab glyph's caret edge. */
export function tabCaret(): Plugin {
  return new Plugin({
    view(view) {
      const ownerDocument = view.dom.ownerDocument;
      const ownerWindow = ownerDocument.defaultView;
      const caret = ownerDocument.createElement("div");
      caret.className = editorClassNames.tabCaret;
      caret.hidden = true;
      caret.setAttribute("aria-hidden", "true");
      ownerDocument.body.append(caret);

      let frame = 0;
      let live = true;
      const requestFrame = ownerWindow?.requestAnimationFrame.bind(ownerWindow);
      const cancelFrame = ownerWindow?.cancelAnimationFrame.bind(ownerWindow);

      const hide = () => {
        caret.hidden = true;
        view.dom.classList.remove(editorClassNames.tabCaretActive);
      };
      const draw = () => {
        frame = 0;
        if (!live || !view.hasFocus()) {
          hide();
          return;
        }
        const boundary = tabBoundary(view);
        const slot = boundary && tabSlot(view, boundary.position);
        if (!boundary || !slot || !slot.hasAttribute("data-tab-layout")) {
          hide();
          return;
        }
        const slotRect = slot.getBoundingClientRect();
        const paragraph = slot.closest(`.${editorClassNames.paragraph}`);
        const style = ownerWindow?.getComputedStyle(paragraph ?? slot);
        const rtl = style?.direction === "rtl";
        const leading = boundary.side === "before";
        const left = leading === !rtl ? slotRect.left : slotRect.right;
        const coordinates = view.coordsAtPos(view.state.selection.head);
        const height = Math.max(1, coordinates.bottom - coordinates.top);

        caret.style.color =
          ownerWindow?.getComputedStyle(slot).color ?? "currentcolor";
        caret.style.left = `${left}px`;
        caret.style.top = `${coordinates.top}px`;
        caret.style.height = `${height}px`;
        caret.hidden = false;
        view.dom.classList.add(editorClassNames.tabCaretActive);
      };
      const schedule = () => {
        if (!live || frame !== 0) return;
        frame = (requestFrame ?? requestAnimationFrame)(draw);
      };
      const onFocus = () => schedule();
      const onBlur = () => hide();
      const onViewportChange = () => schedule();
      const layer = view.dom.closest(`.${editorClassNames.pageLayer}`);
      const root = view.dom.closest(`.${editorClassNames.root}`);
      const workspace = view.dom.closest(`.${editorClassNames.workspace}`);
      const ResizeObserverClass = ownerWindow?.ResizeObserver;
      const resize =
        typeof ResizeObserverClass === "undefined"
          ? null
          : new ResizeObserverClass(schedule);
      resize?.observe(view.dom);
      if (root) resize?.observe(root);
      if (workspace) resize?.observe(workspace);
      const MutationObserverClass = ownerWindow?.MutationObserver;
      const layoutMutation =
        typeof MutationObserverClass === "undefined"
          ? null
          : new MutationObserverClass(schedule);
      if (layer) {
        layoutMutation?.observe(layer, {
          attributes: true,
          attributeFilter: ["class", "style"],
        });
      }
      if (workspace) {
        layoutMutation?.observe(workspace, {
          attributes: true,
          attributeFilter: ["class", "data-comments", "style"],
        });
      }
      view.dom.addEventListener("focus", onFocus);
      view.dom.addEventListener("blur", onBlur);
      ownerDocument.addEventListener("scroll", onViewportChange, true);
      ownerWindow?.addEventListener("resize", onViewportChange);
      schedule();

      return {
        update: () => schedule(),
        destroy() {
          live = false;
          if (frame !== 0) (cancelFrame ?? cancelAnimationFrame)(frame);
          view.dom.removeEventListener("focus", onFocus);
          view.dom.removeEventListener("blur", onBlur);
          ownerDocument.removeEventListener("scroll", onViewportChange, true);
          ownerWindow?.removeEventListener("resize", onViewportChange);
          resize?.disconnect();
          layoutMutation?.disconnect();
          view.dom.classList.remove(editorClassNames.tabCaretActive);
          caret.remove();
        },
      };
    },
  });
}
