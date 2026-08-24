import type { Mark, Node as PMNode } from "prosemirror-model";
import { editRunProps, type RunEdit } from "../../docx/runProps";
import type { RunFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { editorClassNames } from "../../styles/classNames";

const IGNORED_TAGS = new Set([
  "HEAD",
  "IFRAME",
  "NOSCRIPT",
  "OBJECT",
  "SCRIPT",
  "STYLE",
  "TEMPLATE",
]);

const GENERIC_FONTS = new Set([
  "cursive",
  "fantasy",
  "monospace",
  "sans-serif",
  "serif",
  "system-ui",
]);

const NAMED_COLORS: Readonly<Record<string, string>> = {
  aqua: "#00ffff",
  black: "#000000",
  blue: "#0000ff",
  fuchsia: "#ff00ff",
  gray: "#808080",
  green: "#008000",
  lime: "#00ff00",
  maroon: "#800000",
  navy: "#000080",
  olive: "#808000",
  orange: "#ffa500",
  purple: "#800080",
  red: "#ff0000",
  silver: "#c0c0c0",
  teal: "#008080",
  white: "#ffffff",
  yellow: "#ffff00",
};

export interface InlineStyle {
  bold?: true;
  italic?: true;
  underline?: true;
  strike?: true;
  fontSizePt?: number;
  fontFamily?: string;
  color?: string;
  background?: string;
}

export interface InlineContext {
  style: InlineStyle;
  href: string | null;
  preserveWhitespace: boolean;
}

export type InlineElementReader = (
  target: PMNode[],
  element: HTMLElement,
  context: InlineContext
) => boolean;

function cleanText(source: string, preserveWhitespace: boolean): string {
  // biome-ignore lint/suspicious/noControlCharactersInRegex: Word uses these clipboard characters as line breaks
  const withoutForbidden = source.replace(/[\u000B\u000C]/g, "\n").replace(
    // biome-ignore lint/suspicious/noControlCharactersInRegex: XML 1.0 cannot hold these clipboard characters
    /[\u0000-\u0008\u000E-\u001F]/g,
    ""
  );
  return preserveWhitespace
    ? withoutForbidden.replace(/\r\n?/g, "\n")
    : withoutForbidden.replace(/[\t\n\v\f\r ]+/g, " ");
}

function cssHex(value: string): string | null {
  const source = value.trim().toLowerCase();
  const named = NAMED_COLORS[source];
  if (named) return named;
  const short = /^#([0-9a-f]{3})$/.exec(source)?.[1];
  if (short) {
    return `#${short
      .split("")
      .map((part) => `${part}${part}`)
      .join("")}`;
  }
  const full = /^#([0-9a-f]{6})(?:ff)?$/.exec(source)?.[1];
  if (full) return `#${full}`;
  const rgb =
    /^rgba?\(\s*([\d.]+)[, ]+\s*([\d.]+)[, ]+\s*([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?\s*\)$/.exec(
      source
    );
  if (!rgb) return null;
  const alpha = rgb[4];
  if (alpha !== undefined) {
    const opacity = Number.parseFloat(alpha);
    if (alpha.endsWith("%") ? opacity < 100 : opacity < 1) return null;
  }
  const channels = rgb
    .slice(1, 4)
    .map((part) =>
      Math.max(0, Math.min(255, Math.round(Number.parseFloat(part ?? "0"))))
    );
  return `#${channels.map((part) => part.toString(16).padStart(2, "0")).join("")}`;
}

function fontSizePt(value: string): number | null {
  const match = /^([\d.]+)(pt|px)$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number.parseFloat(match[1] ?? "");
  if (!Number.isFinite(amount) || amount <= 0) return null;
  const points = match[2]?.toLowerCase() === "px" ? amount * 0.75 : amount;
  const rounded = Math.round(points * 2) / 2;
  return rounded > 0 && rounded <= 819 ? rounded : null;
}

function legacyFontSize(value: string | null): number | null {
  const sizes = [8, 10, 12, 14, 18, 24, 36];
  const index = Number.parseInt(value ?? "", 10) - 1;
  return sizes[index] ?? null;
}

function firstFontFamily(value: string): string | null {
  const first = value
    .split(",", 1)[0]
    ?.trim()
    .replace(/^(['"])(.*)\1$/, "$2");
  if (!first || GENERIC_FONTS.has(first.toLowerCase())) return null;
  return first;
}

function toggle(
  style: InlineStyle,
  key: "bold" | "italic" | "underline" | "strike",
  on: boolean
): void {
  if (on) style[key] = true;
  else delete style[key];
}

function styleOf(parent: InlineStyle, element: HTMLElement): InlineStyle {
  const style = { ...parent };
  const tag = element.tagName;
  if (tag === "B" || tag === "STRONG") style.bold = true;
  if (tag === "I" || tag === "EM") style.italic = true;
  if (tag === "U") style.underline = true;
  if (tag === "S" || tag === "STRIKE" || tag === "DEL") style.strike = true;
  if (tag === "MARK") style.background = "#ffff00";

  const weight = element.style.fontWeight.trim().toLowerCase();
  if (weight) {
    const numeric = Number.parseInt(weight, 10);
    toggle(
      style,
      "bold",
      weight === "bold" || weight === "bolder" || numeric >= 600
    );
  }
  const fontStyle = element.style.fontStyle.trim().toLowerCase();
  if (fontStyle) {
    toggle(style, "italic", fontStyle === "italic" || fontStyle === "oblique");
  }
  const decoration = (
    element.style.textDecorationLine || element.style.textDecoration
  ).toLowerCase();
  if (decoration) {
    toggle(style, "underline", decoration.includes("underline"));
    toggle(style, "strike", decoration.includes("line-through"));
  }

  const size =
    fontSizePt(element.style.fontSize) ||
    (tag === "FONT" ? legacyFontSize(element.getAttribute("size")) : null);
  if (size !== null) style.fontSizePt = size;
  const family = firstFontFamily(
    element.style.fontFamily ||
      (tag === "FONT" ? (element.getAttribute("face") ?? "") : "")
  );
  if (family !== null) style.fontFamily = family;
  const color = cssHex(
    element.style.color ||
      (tag === "FONT" ? (element.getAttribute("color") ?? "") : "")
  );
  if (color !== null) style.color = color;
  const background = cssHex(element.style.backgroundColor);
  if (background !== null) style.background = background;
  return style;
}

function safeHref(value: string | null): string | null {
  const href = value?.trim() ?? "";
  if (href === "") return null;
  if (/^(?:https?|mailto|tel):/i.test(href)) return href;
  return /^(?:[./]|#)/.test(href) ? href : null;
}

export function contextFor(
  parent: InlineContext,
  element: HTMLElement,
  readElementStyle = true
): InlineContext {
  const copiedEditorLink = element.classList.contains(editorClassNames.link)
    ? element.getAttribute("data-href")
    : null;
  return {
    style: readElementStyle ? styleOf(parent.style, element) : parent.style,
    href:
      element.tagName === "A"
        ? safeHref(element.getAttribute("href"))
        : (safeHref(copiedEditorLink) ?? parent.href),
    preserveWhitespace: parent.preserveWhitespace || element.tagName === "PRE",
  };
}

export function withInlineStyle(
  context: InlineContext,
  style: InlineStyle
): InlineContext {
  return { ...context, style: { ...context.style, ...style } };
}

function runMark(style: InlineStyle): Mark | null {
  const edits: RunEdit[] = [];
  if (style.bold) edits.push({ kind: "toggle", toggle: "bold", on: true });
  if (style.italic) edits.push({ kind: "toggle", toggle: "italic", on: true });
  if (style.underline) {
    edits.push({ kind: "toggle", toggle: "underline", on: true });
  }
  if (style.strike) edits.push({ kind: "toggle", toggle: "strike", on: true });
  if (style.fontSizePt !== undefined) {
    edits.push({ kind: "fontSize", pt: style.fontSizePt });
  }
  if (style.fontFamily !== undefined) {
    edits.push({ kind: "fontFamily", name: style.fontFamily });
  }
  if (style.color !== undefined)
    edits.push({ kind: "color", hex: style.color });
  if (style.background !== undefined) {
    edits.push({ kind: "background", hex: style.background });
  }
  let props: { rPr: string | null; format: RunFormat | null } = {
    rPr: null,
    format: null,
  };
  for (const edit of edits) {
    const next = editRunProps(props, null, edit);
    if (next) props = next;
  }
  return props.rPr === null
    ? null
    : docxSchema.marks.run.create({ rPr: props.rPr, format: props.format });
}

export function marksFor(context: InlineContext): Mark[] {
  const marks: Mark[] = [];
  if (context.href !== null) {
    marks.push(docxSchema.marks.link.create({ href: context.href }));
  }
  const run = runMark(context.style);
  if (run) marks.push(run);
  return marks;
}

function appendText(
  target: PMNode[],
  source: string,
  context: InlineContext
): void {
  const value = cleanText(source, context.preserveWhitespace);
  if (value === "" || (value === " " && target.length === 0)) return;
  const marks = marksFor(context);
  const parts = context.preserveWhitespace ? value.split("\n") : [value];
  parts.forEach((part, index) => {
    if (index > 0) {
      target.push(docxSchema.nodes.hardBreak.create(null, null, marks));
    }
    if (part !== "") target.push(docxSchema.text(part, marks));
  });
}

export function appendInline(
  target: PMNode[],
  node: Node,
  context: InlineContext,
  readElement?: InlineElementReader
): void {
  if (node.nodeType === node.TEXT_NODE) {
    appendText(target, node.nodeValue ?? "", context);
    return;
  }
  if (node.nodeType !== node.ELEMENT_NODE) return;
  const element = node as HTMLElement;
  if (IGNORED_TAGS.has(element.tagName)) return;
  const next = contextFor(context, element);
  if (readElement?.(target, element, next)) return;
  if (element.classList.contains(editorClassNames.tab)) {
    const marks = marksFor(next);
    const tab = docxSchema.marks.tab.create();
    target.push(docxSchema.text("\t", tab.addToSet(marks)));
    return;
  }
  if (element.tagName === "BR") {
    target.push(docxSchema.nodes.hardBreak.create(null, null, marksFor(next)));
    return;
  }
  for (const child of element.childNodes) {
    appendInline(target, child, next, readElement);
  }
}
