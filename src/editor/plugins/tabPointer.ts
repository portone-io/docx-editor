import { Plugin, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { editorClassNames } from "../../styles/classNames";

type TabSide = "before" | "after";

/** Maps the visual half of a tab to its logical side in the text flow. */
export function tabSideAtX(
  left: number,
  right: number,
  clientX: number,
  direction: "ltr" | "rtl"
): TabSide {
  const leftHalf = clientX < (left + right) / 2;
  return leftHalf === (direction === "ltr") ? "before" : "after";
}

function tabSlot(target: EventTarget | null): HTMLElement | null {
  if (!target || typeof (target as Element).closest !== "function") return null;
  return (target as Element).closest<HTMLElement>(
    `.${editorClassNames.tabSlot}`
  );
}

function tabPosition(view: EditorView, slot: HTMLElement, clientX: number) {
  const paragraph = slot.closest(`.${editorClassNames.paragraph}`);
  const direction =
    paragraph &&
    paragraph.ownerDocument.defaultView?.getComputedStyle(paragraph)
      .direction === "rtl"
      ? "rtl"
      : "ltr";
  const rect = slot.getBoundingClientRect();
  const side = tabSideAtX(rect.left, rect.right, clientX, direction);
  return view.posAtDOM(slot, side === "before" ? 0 : slot.childNodes.length);
}

function pointerPosition(view: EditorView, event: MouseEvent): number | null {
  const direct = tabSlot(event.target);
  const ownerDocument = view.dom.ownerDocument;
  const pointed =
    typeof ownerDocument.elementFromPoint === "function"
      ? ownerDocument.elementFromPoint(event.clientX, event.clientY)
      : null;
  const slot = direct ?? tabSlot(pointed);
  if (slot && view.dom.contains(slot)) {
    return tabPosition(view, slot, event.clientX);
  }
  return (
    view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ?? null
  );
}

function select(view: EditorView, anchor: number, head: number): void {
  const size = view.state.doc.content.size;
  const boundedAnchor = Math.max(0, Math.min(anchor, size));
  const boundedHead = Math.max(0, Math.min(head, size));
  const selection = TextSelection.between(
    view.state.doc.resolve(boundedAnchor),
    view.state.doc.resolve(boundedHead)
  );
  view.dispatch(view.state.tr.setSelection(selection));
}

/** Gives a tab deterministic click and drag selection boundaries. */
export function tabPointer(): Plugin {
  let stopDragging: (() => void) | null = null;
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown(view, event) {
          if (event.button !== 0) return false;
          const slot = tabSlot(event.target);
          if (!slot || !view.dom.contains(slot)) return false;

          event.preventDefault();
          view.focus();
          const head = tabPosition(view, slot, event.clientX);
          const anchor = event.shiftKey ? view.state.selection.anchor : head;
          select(view, anchor, head);

          stopDragging?.();
          const onMove = (move: MouseEvent) => {
            const next = pointerPosition(view, move);
            if (next === null) return;
            move.preventDefault();
            select(view, anchor, next);
          };
          const onUp = (up: MouseEvent) => {
            const next = pointerPosition(view, up);
            if (next !== null) select(view, anchor, next);
            stopDragging?.();
          };
          const ownerDocument = view.dom.ownerDocument;
          stopDragging = () => {
            ownerDocument.removeEventListener("mousemove", onMove, true);
            ownerDocument.removeEventListener("mouseup", onUp, true);
            stopDragging = null;
          };
          ownerDocument.addEventListener("mousemove", onMove, true);
          ownerDocument.addEventListener("mouseup", onUp, true);
          return true;
        },
      },
    },
    view: (view) => {
      const ownerWindow = view.dom.ownerDocument.defaultView;
      const onBlur = () => stopDragging?.();
      ownerWindow?.addEventListener("blur", onBlur);
      return {
        destroy() {
          ownerWindow?.removeEventListener("blur", onBlur);
          stopDragging?.();
        },
      };
    },
  });
}
