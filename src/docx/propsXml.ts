/**
 * Handles the XML fragments that carry formatting, such as `<w:rPr>`, `<w:pPr>`, and `<w:tcPr>`.
 *
 * A fragment is never rebuilt: only the one child a job names is sliced out and swapped, so the
 * rest (borders, shading, margins) stays as the original text wrote it. OOXML lays down the order
 * of the children, so the spot to insert a child that was not there is found by that same order.
 */

import { elementChildren, localPart, parseXml, W_NS } from "../ooxml/xml";

export interface PropsChild {
  /** The name with its namespace prefix stripped off (e.g. `gridSpan`) */
  name: string;
  /** This child's original XML fragment exactly as it was */
  xml: string;
}

export interface Props {
  /** The opening tag name exactly as written (e.g. `w:tcPr`) */
  tag: string;
  attrs: string | null;
  children: PropsChild[];
}

/** The order the children are laid out in under `w:rPr` (CT_RPr) */
export const RUN_PR_ORDER: readonly string[] = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
  "rPrChange",
];

/** The order the children are laid out in under `w:pPr` (CT_PPr) */
export const P_PR_ORDER: readonly string[] = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
];

/** The order the children are laid out in under `w:tcPr` (CT_TcPr) */
export const TC_PR_ORDER: readonly string[] = [
  "cnfStyle",
  "tcW",
  "gridSpan",
  "hMerge",
  "vMerge",
  "tcBorders",
  "shd",
  "noWrap",
  "tcMar",
  "textDirection",
  "tcFitText",
  "vAlign",
  "hideMark",
  "headers",
  "cellIns",
  "cellDel",
  "cellMerge",
  "tcPrChange",
];

/** The conventional order of children under `w:trPr` (CT_TrPr). */
export const TR_PR_ORDER: readonly string[] = [
  "cnfStyle",
  "divId",
  "gridBefore",
  "gridAfter",
  "wBefore",
  "wAfter",
  "cantSplit",
  "trHeight",
  "tblHeader",
  "tblCellSpacing",
  "jc",
  "hidden",
  "ins",
  "del",
  "trPrChange",
];

/** The order the children are laid out in under `w:tblPr` (CT_TblPr) */
export const TBL_PR_ORDER: readonly string[] = [
  "tblStyle",
  "tblpPr",
  "tblOverlap",
  "bidiVisual",
  "tblStyleRowBandSize",
  "tblStyleColBandSize",
  "tblW",
  "jc",
  "tblCellSpacing",
  "tblInd",
  "tblBorders",
  "shd",
  "tblLayout",
  "tblCellMar",
  "tblLook",
  "tblCaption",
  "tblDescription",
  "tblPrChange",
];

type TagKind = "open" | "close" | "empty" | "other";

interface Tag {
  kind: TagKind;
  name: string;
  /** The spot where the tag name ends. This is the start of the attribute string */
  nameEnd: number;
  /** The spot right after the tag ends */
  end: number;
}

function skipPast(source: string, from: number, marker: string): number {
  const at = source.indexOf(marker, from);
  return at === -1 ? -1 : at + marker.length;
}

function otherTag(end: number): Tag | null {
  return end === -1 ? null : { kind: "other", name: "", nameEnd: end, end };
}

/** Reads a single tag starting at a `<`. null if it cannot be made out */
function readTag(source: string, lt: number): Tag | null {
  if (source.startsWith("<?", lt)) return otherTag(skipPast(source, lt, "?>"));
  if (source.startsWith("<!--", lt))
    return otherTag(skipPast(source, lt, "-->"));
  if (source.startsWith("<![CDATA[", lt))
    return otherTag(skipPast(source, lt, "]]>"));
  if (source.startsWith("<!", lt)) return otherTag(skipPast(source, lt, ">"));

  const closing = source.startsWith("</", lt);
  let i = lt + (closing ? 2 : 1);
  const nameStart = i;
  while (i < source.length && !" \t\r\n/>".includes(source[i])) i += 1;
  const name = source.slice(nameStart, i);
  if (!name) return null;
  const nameEnd = i;

  let quote: string | null = null;
  while (i < source.length) {
    const ch = source[i];
    if (quote !== null) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      const selfClosing = source[i - 1] === "/";
      const kind: TagKind = closing ? "close" : selfClosing ? "empty" : "open";
      return { kind, name, nameEnd, end: i + 1 };
    }
    i += 1;
  }
  return null;
}

function attrsOf(source: string, tag: Tag): string | null {
  const closeLength = tag.kind === "empty" ? 2 : 1;
  const attrs = source.slice(tag.nameEnd, tag.end - closeLength).trim();
  return attrs.length > 0 ? attrs : null;
}

/**
 * Splits a formatting fragment into its opening tag and its list of children.
 * null if its shape cannot be made out (in which case leaving the original untouched is the safe move).
 */
export function parseProps(xml: string): Props | null {
  const open = readTag(xml, 0);
  if (!open || xml[0] !== "<") return null;
  const attrs = attrsOf(xml, open);
  if (open.kind === "empty") {
    return open.end === xml.length
      ? { tag: open.name, attrs, children: [] }
      : null;
  }
  if (open.kind !== "open") return null;

  const children: PropsChild[] = [];
  let depth = 0;
  let childStart = -1;
  let childName = "";
  let i = open.end;

  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    if (lt === -1) return null;
    const tag = readTag(xml, lt);
    if (!tag) return null;

    if (tag.kind === "other") {
      i = tag.end;
      continue;
    }

    if (depth === 0) {
      if (tag.kind === "close") {
        if (tag.name !== open.name || tag.end !== xml.length) return null;
        return { tag: open.name, attrs, children };
      }
      childStart = lt;
      childName = tag.name;
      if (tag.kind === "empty") {
        children.push({
          name: localPart(childName),
          xml: xml.slice(childStart, tag.end),
        });
      } else {
        depth = 1;
      }
    } else if (tag.kind === "open") {
      depth += 1;
    } else if (tag.kind === "close") {
      depth -= 1;
      if (depth === 0) {
        children.push({
          name: localPart(childName),
          xml: xml.slice(childStart, tag.end),
        });
      }
    }
    i = tag.end;
  }
  return null;
}

/** With no children at all, the formatting fragment itself is not written */
export function renderProps(props: Props): string {
  if (props.children.length === 0) return "";
  const open = props.attrs ? `<${props.tag} ${props.attrs}>` : `<${props.tag}>`;
  return (
    open + props.children.map((child) => child.xml).join("") + `</${props.tag}>`
  );
}

/**
 * The spot this child goes into within the prescribed order.
 * A child whose place in the order we do not know is treated as sitting right behind the child before it.
 */
function insertIndex(
  children: PropsChild[],
  name: string,
  order: readonly string[]
): number {
  const target = order.indexOf(name);
  if (target === -1) return children.length;
  let previous = -1;
  for (const [index, child] of children.entries()) {
    const known = order.indexOf(child.name);
    const effective = known === -1 ? previous : known;
    if (effective > target) return index;
    previous = effective;
  }
  return children.length;
}

/**
 * Replaces a single child with new XML.
 * A null `xml` removes that child. A child that was not there goes into the spot the order calls for.
 */
export function setPropsChild(
  children: PropsChild[],
  name: string,
  xml: string | null,
  order: readonly string[]
): PropsChild[] {
  const without = children.filter((child) => child.name !== name);
  if (xml === null) return without;

  const child: PropsChild = { name, xml };
  const at = children.findIndex((entry) => entry.name === name);
  if (at === -1) {
    const index = insertIndex(without, name, order);
    return [...without.slice(0, index), child, ...without.slice(index)];
  }
  // Keeps the spot it originally occupied. If the same name appears several times, only the first spot survives
  return [
    ...children.slice(0, at),
    child,
    ...children.slice(at + 1).filter((entry) => entry.name !== name),
  ];
}

export function propsChild(
  children: PropsChild[],
  name: string
): PropsChild | undefined {
  return children.find((child) => child.name === name);
}

/**
 * Gathers the namespace prefixes used in the fragment and declares them.
 * Only `w` carries real meaning; the rest are placeholders that keep the parser from stopping.
 * All we read are element names and `w:` attributes, so placeholders still let the values be read as they are.
 */
function namespaceDecls(xml: string): string {
  const prefixes = new Set<string>(["w"]);
  for (const [, prefix] of xml.matchAll(/<\/?([A-Za-z_][\w.-]*):/g)) {
    prefixes.add(prefix);
  }
  for (const [, prefix] of xml.matchAll(/[\s"']([A-Za-z_][\w.-]*):[\w.-]+=/g)) {
    prefixes.add(prefix);
  }
  // `xml` and `xmlns` are names that cannot be redeclared. Declaring them makes parsing fail
  prefixes.delete("xml");
  prefixes.delete("xmlns");
  return Array.from(prefixes)
    .map(
      (prefix) =>
        `xmlns:${prefix}="${prefix === "w" ? W_NS : `urn:docx-editor:${prefix}`}"`
    )
    .join(" ");
}

/**
 * Reads a single formatting fragment into an element.
 *
 * A fragment carries no namespace declarations, so they are put back on as it is wrapped.
 * null if its shape cannot be made out, in which case the caller leaves the display values alone.
 */
export function parsePropsXml(xml: string): Element | null {
  try {
    const wrapped = `<props ${namespaceDecls(xml)}>${xml}</props>`;
    return elementChildren(parseXml(wrapped).documentElement)[0] ?? null;
  } catch {
    return null;
  }
}
