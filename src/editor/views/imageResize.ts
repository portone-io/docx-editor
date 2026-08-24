/**
 * Image node view with aspect-ratio-locked corner resizing. Dragging previews in the DOM and
 * commits one EMU transaction on release so undo history receives a single step.
 */

import { DOMSerializer, type Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView, NodeView } from "prosemirror-view";
import { emuToPx, pxToEmu, toImageExtent } from "../../ooxml/image";
import { docxSchema, imageNodeSpec } from "../../schema";
import { editorAttributes, editorClassNames } from "../../styles/classNames";

/** The corners an image can be grabbed by */
const CORNERS = ["nw", "ne", "sw", "se"] as const;

export type ImageCorner = (typeof CORNERS)[number];

/** A size on screen, in pixels */
export interface ImageSizePx {
  width: number;
  height: number;
}

/** The smallest an image may be dragged down to. Below this there is nothing left to grab */
const MIN_IMAGE_PX = 16;

/**
 * The size a corner drag asks for, with the ratio locked.
 *
 * The two distances are turned into relative ones and averaged into a single scale. That
 * way a drag along the image's own diagonal follows the pointer exactly, and a drag along
 * one axis alone still resizes.
 */
export function resizedImagePx(
  start: ImageSizePx,
  corner: ImageCorner,
  dx: number,
  dy: number
): ImageSizePx {
  const sideways =
    (corner === "ne" || corner === "se" ? dx : -dx) / start.width;
  const downward =
    (corner === "sw" || corner === "se" ? dy : -dy) / start.height;
  const scale = Math.max(
    1 + (sideways + downward) / 2,
    MIN_IMAGE_PX / start.width,
    MIN_IMAGE_PX / start.height
  );
  return {
    width: Math.round(start.width * scale),
    height: Math.round(start.height * scale),
  };
}

/**
 * Records a new display size on the image at this position.
 * Null when there is no image there, or when the size is the one it already holds, so a
 * drag that moved nothing leaves no step in the history.
 */
export function buildResizeImageTransaction(
  state: EditorState,
  pos: number,
  size: ImageSizePx
): Transaction | null {
  const node = state.doc.nodeAt(pos);
  if (node?.type !== docxSchema.nodes.image) return null;
  const extent = { cx: pxToEmu(size.width), cy: pxToEmu(size.height) };
  const current = toImageExtent(node.attrs.extent);
  if (current?.cx === extent.cx && current.cy === extent.cy) return null;
  return state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, extent });
}

/** The `img` element the schema draws an image node as */
function renderImage(node: PMNode): HTMLImageElement {
  const { dom } = DOMSerializer.renderSpec(document, imageNodeSpec(node.attrs));
  if (!(dom instanceof HTMLImageElement)) {
    throw new Error("the image was not rendered as an img element");
  }
  return dom;
}

/**
 * Brings the drawn element back in line with the node, attribute by attribute.
 *
 * The element itself lives on, so the bytes it has already decoded are not thrown away
 * every time the size changes.
 */
function syncAttributes(img: HTMLImageElement, node: PMNode): void {
  const next = renderImage(node);
  for (const name of img.getAttributeNames()) {
    if (!next.hasAttribute(name)) img.removeAttribute(name);
  }
  for (const name of next.getAttributeNames()) {
    const value = next.getAttribute(name) ?? "";
    if (img.getAttribute(name) !== value) img.setAttribute(name, value);
  }
}

function handleElement(corner: ImageCorner): HTMLElement {
  const handle = document.createElement("span");
  handle.className = editorClassNames.imageHandle;
  handle.setAttribute(editorAttributes.imageCorner, corner);
  return handle;
}

/** The corner handle an event happened on. Null for anywhere else on the image */
function cornerAt(target: EventTarget | null): ImageCorner | null {
  const element = target instanceof Element ? target : null;
  const corner = element?.getAttribute(editorAttributes.imageCorner);
  return CORNERS.find((one) => one === corner) ?? null;
}

/** Draws the image at this size, leaving the document as it is */
function showSize(img: HTMLImageElement, size: ImageSizePx): void {
  img.width = size.width;
  img.height = size.height;
}

/** What is held on to while a corner is being dragged */
interface ImageDrag {
  readonly corner: ImageCorner;
  readonly startX: number;
  readonly startY: number;
  /** The size the image stood at when the drag started */
  readonly start: ImageSizePx;
  /** The size the drag is currently asking for */
  size: ImageSizePx;
}

export class ImageNodeView implements NodeView {
  readonly dom: HTMLElement;
  private readonly view: EditorView;
  private readonly getPos: () => number | undefined;
  private readonly img: HTMLImageElement;
  private node: PMNode;
  private drag: ImageDrag | null = null;

  constructor(
    node: PMNode,
    view: EditorView,
    getPos: () => number | undefined
  ) {
    this.node = node;
    this.view = view;
    this.getPos = getPos;
    this.img = renderImage(node);
    this.dom = document.createElement("span");
    this.dom.className = editorClassNames.imageBox;
    this.dom.append(this.img, ...CORNERS.map(handleElement));
    this.dom.addEventListener("mousedown", this.onMouseDown);
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    syncAttributes(this.img, node);
    return true;
  }

  /** Only an editable image shows its handles: there is nothing to drag in a read-only document */
  selectNode(): void {
    if (this.view.editable) {
      this.dom.classList.add(editorClassNames.imageSelected);
    }
  }

  deselectNode(): void {
    this.dom.classList.remove(editorClassNames.imageSelected);
  }

  /** A press on a handle is ours alone. It must not move the selection or start a drag of the image */
  stopEvent(event: Event): boolean {
    return cornerAt(event.target) !== null;
  }

  destroy(): void {
    this.stopDrag();
    this.dom.removeEventListener("mousedown", this.onMouseDown);
  }

  /** The size the image stands at now: the one the document records, or the drawn size when it records none */
  private currentSizePx(): ImageSizePx | null {
    const extent = toImageExtent(this.node.attrs.extent);
    if (extent) {
      return { width: emuToPx(extent.cx), height: emuToPx(extent.cy) };
    }
    const rect = this.img.getBoundingClientRect();
    if (!(rect.width > 0) || !(rect.height > 0)) return null;
    return { width: rect.width, height: rect.height };
  }

  private readonly onMouseDown = (event: MouseEvent): void => {
    const corner = cornerAt(event.target);
    if (corner === null || this.drag || event.button !== 0) return;
    if (!this.view.editable) return;
    const start = this.currentSizePx();
    if (!start) return;

    // A press on a handle starts neither a text selection nor a drag of the image
    event.preventDefault();
    this.drag = {
      corner,
      startX: event.clientX,
      startY: event.clientY,
      start,
      size: start,
    };
    window.addEventListener("mousemove", this.onMove, true);
    window.addEventListener("mouseup", this.onUp, true);
    window.addEventListener("keydown", this.onKey, true);
  };

  private readonly onMove = (event: MouseEvent): void => {
    const drag = this.drag;
    if (!drag) return;
    event.preventDefault();
    drag.size = resizedImagePx(
      drag.start,
      drag.corner,
      event.clientX - drag.startX,
      event.clientY - drag.startY
    );
    showSize(this.img, drag.size);
  };

  private readonly onUp = (): void => {
    const size = this.drag?.size;
    const pos = this.getPos();
    // The commit redraws the image, so the preview is taken down first
    this.stopDrag();
    if (!size || pos === undefined) return;
    const tr = buildResizeImageTransaction(this.view.state, pos, size);
    if (tr) this.view.dispatch(tr);
  };

  private readonly onKey = (event: KeyboardEvent): void => {
    if (event.key !== "Escape") return;
    event.preventDefault();
    this.stopDrag();
  };

  private stopDrag(): void {
    if (!this.drag) return;
    window.removeEventListener("mousemove", this.onMove, true);
    window.removeEventListener("mouseup", this.onUp, true);
    window.removeEventListener("keydown", this.onKey, true);
    this.drag = null;
    // The preview touched nothing but the drawn size, so the node's own size draws over it
    syncAttributes(this.img, this.node);
  }
}
