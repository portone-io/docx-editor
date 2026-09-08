import type { Attrs, DOMOutputSpec } from "prosemirror-model";
import { toRunFormat } from "../model/format";
import { emuToPx, toImageExtent, toImageSrc } from "../ooxml/image";
import { editorClassNames } from "../styles/classNames";
import type { FontFallbacks } from "../styles/fontStack";
import { runStyle } from "../styles/inlineStyle";

function text(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function formatJson(format: object | null): string | undefined {
  return format ? JSON.stringify(format) : undefined;
}

/**
 * What a run says to anyone reading it, the page and a copy of it alike.
 *
 * The rest of the drawing is this editor talking to itself - the original XML, the worked-out
 * formatting it was drawn from - and `schema/clipboard` publishes only what is here.
 */
export function runAttrs(
  attrs: Attrs,
  fontFallbacks: FontFallbacks
): Record<string, string | undefined> {
  const format = toRunFormat(attrs.format);
  return {
    class: editorClassNames.run,
    style: runStyle(format, fontFallbacks),
    // Which shape of a Han character the browser draws is decided by this and nothing else
    lang: format?.lang,
  };
}

/** The DOM representation shared by the schema and the editor's custom run mark view. */
export function runMarkSpec(
  attrs: Attrs,
  fontFallbacks: FontFallbacks
): DOMOutputSpec {
  const format = toRunFormat(attrs.format);
  return [
    "span",
    {
      ...runAttrs(attrs, fontFallbacks),
      "data-rattrs": text(attrs.rAttrs),
      "data-rpr": text(attrs.rPr),
      "data-fmt": formatJson(format),
      // Read by `styles/editor.css` to stand a link's own line down
      "data-underline":
        format?.underline && format.underline !== "none" ? "" : undefined,
    },
    0,
  ];
}

/**
 * What an image says to anyone reading it, the page and a copy of it alike.
 *
 * The size is given in the pixels it is drawn at rather than as the document's own measure in EMU,
 * which is what makes it the half another application can read.
 */
export function imageAttrs(attrs: Attrs): Record<string, string | undefined> {
  const extent = toImageExtent(attrs.extent);
  return {
    class: editorClassNames.image,
    src: toImageSrc(attrs.src) ?? undefined,
    alt: text(attrs.alt) ?? "",
    width: extent ? `${Math.round(emuToPx(extent.cx))}` : undefined,
    height: extent ? `${Math.round(emuToPx(extent.cy))}` : undefined,
  };
}

/** The image element shared by the schema and the resizable image node view. */
export function imageNodeSpec(attrs: Attrs): DOMOutputSpec {
  return [
    "img",
    {
      ...imageAttrs(attrs),
      "data-extent": formatJson(toImageExtent(attrs.extent)),
      "data-xml": text(attrs.xml),
    },
  ];
}
