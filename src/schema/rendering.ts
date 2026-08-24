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

/** The DOM representation shared by the schema and the editor's custom run mark view. */
export function runMarkSpec(
  attrs: Attrs,
  fontFallbacks: FontFallbacks
): DOMOutputSpec {
  const format = toRunFormat(attrs.format);
  return [
    "span",
    {
      class: editorClassNames.run,
      style: runStyle(format, fontFallbacks),
      // Which shape of a Han character the browser draws is decided by this and nothing else
      lang: format?.lang,
      "data-rattrs": text(attrs.rAttrs),
      "data-rpr": text(attrs.rPr),
      "data-fmt": formatJson(format),
    },
    0,
  ];
}

/** The image element shared by the schema and the resizable image node view. */
export function imageNodeSpec(attrs: Attrs): DOMOutputSpec {
  const extent = toImageExtent(attrs.extent);
  return [
    "img",
    {
      class: editorClassNames.image,
      src: toImageSrc(attrs.src) ?? undefined,
      alt: text(attrs.alt) ?? "",
      width: extent ? `${Math.round(emuToPx(extent.cx))}` : undefined,
      height: extent ? `${Math.round(emuToPx(extent.cy))}` : undefined,
      "data-extent": formatJson(extent),
      "data-xml": text(attrs.xml),
    },
  ];
}
