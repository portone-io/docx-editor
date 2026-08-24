import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { DEFAULT_TAB_STOP_PT } from "../../docx/documentSettings";
import { type ParagraphFormat, toParagraphFormat } from "../../model/format";
import type { TabAlignment, TabStop } from "../../model/tabStops";
import { docxSchema } from "../../schema";
import { editorClassNames, editorCssVariables } from "../../styles/classNames";
import { documentDefaultTabStopPt } from "../documentStyles";
import { setTabWidths, tabWidths } from "./tabDecorations";

const PX_PER_PT = 96 / 72;
const POSITION_EPSILON_PT = 0.01;
const WIDTH_EPSILON_PX = 0.1;
const NATURAL_WIDTH_PX = 100_000;
const MAX_LAYOUT_PASSES = 4;

type PositionedAlignment = Exclude<TabAlignment, "num">;

interface ParagraphLayoutContext {
  format: ParagraphFormat;
  node: PMNode;
}

interface CachedParagraphLayout {
  defaultStopPt: number;
  generation: number;
  paragraphWidth: number;
  scale: number;
  widths: readonly number[];
}

export interface TabTarget {
  align: PositionedAlignment;
  targetPt: number;
}

export function nextTabTarget(
  currentPt: number,
  stops: readonly TabStop[],
  defaultStopPt: number
): TabTarget {
  const custom = stops.find(
    (stop) =>
      stop.align !== "bar" && stop.positionPt > currentPt + POSITION_EPSILON_PT
  );
  if (custom && custom.align !== "bar") {
    return {
      align: custom.align === "num" ? "start" : custom.align,
      targetPt: custom.positionPt,
    };
  }
  const interval = defaultStopPt > 0 ? defaultStopPt : DEFAULT_TAB_STOP_PT;
  return {
    align: "start",
    targetPt: (Math.floor(currentPt / interval) + 1) * interval,
  };
}

export interface TabSegmentMetrics {
  wholePt: number;
  decimalPrefixPt: number;
}

function alignedPrefix(
  align: PositionedAlignment,
  segment: TabSegmentMetrics
): number {
  if (align === "center") return segment.wholePt / 2;
  if (align === "end") return segment.wholePt;
  if (align === "decimal") return segment.decimalPrefixPt;
  return 0;
}

export function tabAdvancePt(
  currentPt: number,
  target: TabTarget,
  segment: TabSegmentMetrics
): number {
  return Math.max(
    0,
    target.targetPt - currentPt - alignedPrefix(target.align, segment)
  );
}

function numeric(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function computedStyle(element: Element): CSSStyleDeclaration {
  return (
    element.ownerDocument.defaultView?.getComputedStyle(element) ??
    getComputedStyle(element)
  );
}

function scaleOf(view: EditorView): number {
  const layer = view.dom.closest(`.${editorClassNames.pageLayer}`);
  const parsed = Number.parseFloat(layer ? computedStyle(layer).zoom : "");
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
}

function paragraphLayoutContext(
  view: EditorView,
  slot: HTMLElement
): ParagraphLayoutContext | null {
  const $pos = view.state.doc.resolve(view.posAtDOM(slot, 0));
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth);
    if (node.type === docxSchema.nodes.paragraph) {
      return { format: toParagraphFormat(node.attrs.format) ?? {}, node };
    }
  }
  return null;
}

function resetSlot(slot: HTMLElement): void {
  slot.removeAttribute("data-tab-layout");
  slot.style.removeProperty(editorCssVariables.tabWidth);
}

function cloneParagraph(
  paragraph: HTMLElement,
  widthPx: number,
  resetTabs: boolean
): HTMLElement {
  const clone = paragraph.cloneNode(true) as HTMLElement;
  const computed = computedStyle(paragraph);
  for (let index = 0; index < computed.length; index += 1) {
    const property = computed.item(index);
    clone.style.setProperty(
      property,
      computed.getPropertyValue(property),
      computed.getPropertyPriority(property)
    );
  }
  clone.style.position = "fixed";
  clone.style.inset = "0 auto auto -200000px";
  clone.style.boxSizing = "border-box";
  clone.style.width = `${widthPx}px`;
  clone.style.maxWidth = "none";
  clone.style.textAlign = "start";
  clone.style.visibility = "hidden";
  clone.style.pointerEvents = "none";
  clone.style.zoom = "1";
  clone.setAttribute("aria-hidden", "true");
  if (resetTabs) {
    for (const slot of clone.querySelectorAll<HTMLElement>(
      `.${editorClassNames.tabSlot}`
    )) {
      resetSlot(slot);
    }
  }
  return clone;
}

function rangeAfter(slot: HTMLElement, paragraph: HTMLElement): Range {
  const boundaries = Array.from(
    paragraph.querySelectorAll<HTMLElement>(`.${editorClassNames.tabSlot}, br`)
  );
  const next = boundaries[boundaries.indexOf(slot) + 1];
  const range = paragraph.ownerDocument.createRange();
  range.setStartAfter(slot);
  if (next) range.setEndBefore(next);
  else range.setEnd(paragraph, paragraph.childNodes.length);
  return range;
}

function rangeWidth(range: Range): number {
  const rects = Array.from(range.getClientRects()).filter(
    (rect) => rect.width > 0 || rect.height > 0
  );
  if (rects.length === 0) return 0;
  return (
    Math.max(...rects.map((rect) => rect.right)) -
    Math.min(...rects.map((rect) => rect.left))
  );
}

function decimalSeparator(node: Node, paragraph: HTMLElement): string {
  const language =
    node.parentElement?.closest<HTMLElement>("[lang]")?.lang ||
    paragraph.lang ||
    paragraph.ownerDocument.documentElement.lang ||
    undefined;
  try {
    const parts = new Intl.NumberFormat(language).formatToParts(1.1);
    return parts.find((part) => part.type === "decimal")?.value ?? ".";
  } catch {
    return ".";
  }
}

function decimalPrefix(range: Range, paragraph: HTMLElement): Range | null {
  const walker = paragraph.ownerDocument.createTreeWalker(
    paragraph,
    paragraph.ownerDocument.defaultView?.NodeFilter.SHOW_TEXT ?? 4
  );
  let node = walker.nextNode();
  while (node) {
    if (range.intersectsNode(node)) {
      const at =
        node.textContent?.indexOf(decimalSeparator(node, paragraph)) ?? -1;
      if (at >= 0) {
        const prefix = range.cloneRange();
        prefix.setEnd(node, at);
        return prefix;
      }
    }
    node = walker.nextNode();
  }
  return null;
}

function segmentMetrics(
  slot: HTMLElement,
  paragraph: HTMLElement
): TabSegmentMetrics {
  const whole = rangeAfter(slot, paragraph);
  return {
    wholePt: rangeWidth(whole) / PX_PER_PT,
    decimalPrefixPt:
      rangeWidth(decimalPrefix(whole, paragraph) ?? whole) / PX_PER_PT,
  };
}

function currentPositionPt(paragraph: HTMLElement, slot: HTMLElement): number {
  const paragraphRect = paragraph.getBoundingClientRect();
  const slotRect = slot.getBoundingClientRect();
  const style = computedStyle(paragraph);
  const rtl = style.direction === "rtl";
  const margin = numeric(rtl ? style.marginRight : style.marginLeft);
  const origin = rtl
    ? paragraphRect.right + margin
    : paragraphRect.left - margin;
  return (rtl ? origin - slotRect.right : slotRect.left - origin) / PX_PER_PT;
}

function applyWidth(slot: HTMLElement, widthPx: number): void {
  slot.setAttribute("data-tab-layout", "");
  slot.style.setProperty(editorCssVariables.tabWidth, `${widthPx}px`);
}

function slotPosition(slot: HTMLElement): number | null {
  const parsed = Number.parseInt(slot.dataset.tabPosition ?? "", 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function measureParagraph(
  paragraph: HTMLElement,
  actualSlots: readonly HTMLElement[],
  format: ParagraphFormat,
  defaultStopPt: number,
  scale: number,
  widths: Map<number, number>
): void {
  const paragraphWidth = paragraph.getBoundingClientRect().width / scale;
  if (paragraphWidth <= 0) return;
  const measured = cloneParagraph(paragraph, paragraphWidth, false);
  const natural = cloneParagraph(paragraph, NATURAL_WIDTH_PX, true);
  paragraph.ownerDocument.body.append(measured, natural);
  try {
    const measuredSlots = Array.from(
      measured.querySelectorAll<HTMLElement>(`.${editorClassNames.tabSlot}`)
    );
    const naturalSlots = Array.from(
      natural.querySelectorAll<HTMLElement>(`.${editorClassNames.tabSlot}`)
    );
    // A new width can wrap this tab or a later one, so remeasure to a bounded fixed point.
    for (let pass = 0; pass < MAX_LAYOUT_PASSES; pass += 1) {
      let changed = false;
      measuredSlots.forEach((slot, index) => {
        const actual = actualSlots[index];
        const naturalSlot = naturalSlots[index];
        if (!actual || !naturalSlot) return;
        const position = slotPosition(actual);
        if (position === null) return;
        const currentPt = currentPositionPt(measured, slot);
        const segment = segmentMetrics(naturalSlot, natural);
        const target = nextTabTarget(
          currentPt,
          format.tabStops ?? [],
          defaultStopPt
        );
        const widthPx = tabAdvancePt(currentPt, target, segment) * PX_PER_PT;
        const previous = numeric(
          slot.style.getPropertyValue(editorCssVariables.tabWidth)
        );
        if (
          !slot.hasAttribute("data-tab-layout") ||
          Math.abs(previous - widthPx) > WIDTH_EPSILON_PX
        ) {
          changed = true;
        }
        widths.set(position, widthPx);
        applyWidth(slot, widthPx);
      });
      if (!changed) break;
    }
  } finally {
    measured.remove();
    natural.remove();
  }
}

function sameWidths(
  left: ReadonlyMap<number, number>,
  right: ReadonlyMap<number, number>
): boolean {
  if (left.size !== right.size) return false;
  for (const [position, width] of left) {
    const other = right.get(position);
    if (other === undefined || Math.abs(other - width) > WIDTH_EPSILON_PX) {
      return false;
    }
  }
  return true;
}

function measureTabs(
  view: EditorView,
  cache: WeakMap<PMNode, CachedParagraphLayout>,
  generation: number
): ReadonlyMap<number, number> {
  const scale = scaleOf(view);
  const defaultStopPt = documentDefaultTabStopPt(view.state);
  const widths = new Map(tabWidths(view.state));
  const paragraphs = new Map<HTMLElement, HTMLElement[]>();
  for (const slot of view.dom.querySelectorAll<HTMLElement>(
    `.${editorClassNames.tabSlot}`
  )) {
    const closest = slot.closest(`.${editorClassNames.paragraph}`);
    if (!closest) continue;
    const paragraph = closest as HTMLElement;
    const entries = paragraphs.get(paragraph);
    if (entries) entries.push(slot);
    else paragraphs.set(paragraph, [slot]);
  }

  for (const [paragraph, slots] of paragraphs) {
    const context = paragraphLayoutContext(view, slots[0]);
    if (!context) continue;
    const paragraphWidth = paragraph.getBoundingClientRect().width / scale;
    const cached = cache.get(context.node);
    if (
      cached &&
      cached.generation === generation &&
      cached.defaultStopPt === defaultStopPt &&
      cached.scale === scale &&
      Math.abs(cached.paragraphWidth - paragraphWidth) <= WIDTH_EPSILON_PX &&
      cached.widths.length === slots.length
    ) {
      slots.forEach((slot, index) => {
        const position = slotPosition(slot);
        const width = cached.widths[index];
        if (position !== null && width !== undefined)
          widths.set(position, width);
      });
      continue;
    }
    measureParagraph(
      paragraph,
      slots,
      context.format,
      defaultStopPt,
      scale,
      widths
    );
    cache.set(context.node, {
      defaultStopPt,
      generation,
      paragraphWidth,
      scale,
      widths: slots.map((slot) => {
        const position = slotPosition(slot);
        return position === null ? 0 : (widths.get(position) ?? 0);
      }),
    });
  }
  return widths;
}

/** Resolves tab widths against the paragraph's OOXML coordinate system. */
export function tabLayout(): Plugin {
  return new Plugin({
    view(view) {
      let frame = 0;
      let live = true;
      let generation = 0;
      const cache = new WeakMap<PMNode, CachedParagraphLayout>();
      const ownerWindow = view.dom.ownerDocument.defaultView;
      const requestFrame = ownerWindow?.requestAnimationFrame.bind(ownerWindow);
      const cancelFrame = ownerWindow?.cancelAnimationFrame.bind(ownerWindow);
      const schedule = () => {
        if (!live || frame !== 0) return;
        frame = (requestFrame ?? requestAnimationFrame)(() => {
          frame = 0;
          const widths = measureTabs(view, cache, generation);
          if (sameWidths(widths, tabWidths(view.state))) return;
          view.dispatch(
            setTabWidths(view.state.tr, widths).setMeta("addToHistory", false)
          );
        });
      };

      const ResizeObserverClass = ownerWindow?.ResizeObserver;
      const resize =
        typeof ResizeObserverClass === "undefined"
          ? null
          : new ResizeObserverClass(schedule);
      resize?.observe(view.dom);
      const onResize = () => schedule();
      ownerWindow?.addEventListener("resize", onResize);
      const fonts = view.dom.ownerDocument.fonts;
      const onFontsChanged = () => {
        generation += 1;
        schedule();
      };
      fonts?.addEventListener("loadingdone", onFontsChanged);
      fonts?.addEventListener("loadingerror", onFontsChanged);
      fonts?.ready.then(onFontsChanged);
      schedule();

      return {
        update(next, previous) {
          if (!next.state.doc.eq(previous.doc)) schedule();
        },
        destroy() {
          live = false;
          if (frame !== 0) (cancelFrame ?? cancelAnimationFrame)(frame);
          resize?.disconnect();
          ownerWindow?.removeEventListener("resize", onResize);
          fonts?.removeEventListener("loadingdone", onFontsChanged);
          fonts?.removeEventListener("loadingerror", onFontsChanged);
        },
      };
    },
  });
}
