/**
 * Reading the `w:sdt` content controls a document wraps its settled parts in, and naming a
 * control that has to go out more than once.
 *
 * Only the shape we can write back out unchanged is taken apart: the `w:` prefix on both tags
 * we write ourselves, a `w:sdtPr`, and a `w:sdtContent` that carries no attribute of its own
 * and stands last. Everything else is null, which leaves whatever held the control preserved
 * together with its original fragment.
 *
 * What sits inside the control is the caller's business. A whole cell (`docx/importTable`) and
 * a stretch of runs inside a paragraph (`docx/importParagraph`) each read the content their
 * own way, and the wrapper they put back on export is the same string in both cases.
 */

import { wAttr } from "../ooxml/units";
import { attrString, elementChildren, serializeXml } from "../ooxml/xml";
import {
  type Props,
  parseProps,
  propsChild,
  renderProps,
  setPropsChild,
} from "./propsXml";

/** The opening of a content control taken apart, and the content it wraps */
export interface SdtWrapper {
  /** The `<w:sdt>` opening tag followed by everything that stood ahead of `<w:sdtContent>` */
  prefix: string;
  content: Element;
  /** Whether the control says its contents may not be edited */
  contentsLocked: boolean;
  /** Whether the control says it may not be deleted, not even whole */
  deletionLocked: boolean;
}

/**
 * The `w:lock` values that shut each of the two clauses a lock settles (§17.18.49). The two are
 * independent, so a control may be un-editable yet removable (`contentLocked`) or editable yet
 * not removable (`sdtLocked`).
 */
const CONTENTS_LOCKED = ["contentLocked", "sdtContentLocked"];
const DELETION_LOCKED = ["sdtLocked", "sdtContentLocked"];

function lockValue(sdtPr: Element): string | null {
  const lock = elementChildren(sdtPr).find(
    (child) => child.localName === "lock"
  );
  return lock ? wAttr(lock, "val") : null;
}

/** Takes a `w:sdt` apart into the wrapper to put back on export and the content to read. null for a shape we do not write ourselves */
export function readSdtWrapper(el: Element): SdtWrapper | null {
  if (el.nodeName !== "w:sdt") return null;
  const children = elementChildren(el);
  const contentAt = children.findIndex(
    (child) => child.localName === "sdtContent"
  );
  const content = contentAt === -1 ? null : children[contentAt];
  // We write the content tag ourselves, so nothing may hang off it and nothing may follow it
  if (!content || content.nodeName !== "w:sdtContent") return null;
  if (content.attributes.length > 0) return null;
  if (contentAt !== children.length - 1) return null;

  const head = children.slice(0, contentAt);
  const sdtPr = head.find((child) => child.localName === "sdtPr");
  if (!sdtPr) return null;

  const attrs = attrString(el);
  const val = lockValue(sdtPr);
  return {
    prefix:
      (attrs ? `<w:sdt ${attrs}>` : "<w:sdt>") +
      head.map(serializeXml).join(""),
    content,
    contentsLocked: val !== null && CONTENTS_LOCKED.includes(val),
    deletionLocked: val !== null && DELETION_LOCKED.includes(val),
  };
}

/** The order the children are laid out in under `w:sdt` (CT_SdtBlock) */
const SDT_ORDER: readonly string[] = ["sdtPr", "sdtEndPr", "sdtContent"];

/** The order the children are laid out in under `w:sdtPr` (CT_SdtPr) */
const SDT_PR_ORDER: readonly string[] = [
  "rPr",
  "alias",
  "tag",
  "id",
  "lock",
  "placeholder",
  "showingPlcHdr",
  "dataBinding",
  "temporary",
];

/**
 * The number Word writes on every control. Nothing reads it back and it only has to differ
 * from the other controls in the document, so a draw out of the whole 32 bit range is enough.
 */
export function newControlId(): number {
  return Math.floor(Math.random() * 0x7fffffff);
}

function renderPrefix(props: Props): string {
  const open = props.attrs ? `<${props.tag} ${props.attrs}>` : `<${props.tag}>`;
  return open + props.children.map((child) => child.xml).join("");
}

/** What text one child of the control's `w:sdtPr` is to be changed to. A null xml removes that child */
export type SdtPrEdit = readonly [name: string, xml: string | null];

/** What a control carries when it says nothing but which control it is and whether it is shut */
const NAMES_NOTHING: readonly string[] = ["id", "lock"];

/**
 * Whether the control says nothing about itself beyond its id and its lock: no alias or tag, no
 * `w:dataBinding`, no type of its own.
 * A lock may take such a control over a wider stretch, since the stretch is all it ever meant. One
 * Word named keeps the stretch it was given, or the name would come to cover other text.
 */
export function namesNothing(prefix: string): boolean {
  const sdt = parseProps(`${prefix}</w:sdt>`);
  if (!sdt || sdt.attrs !== null) return false;
  const sdtPr = propsChild(sdt.children, "sdtPr");
  const props = sdtPr ? parseProps(sdtPr.xml) : null;
  if (!props || sdt.children.length !== 1) return false;
  return props.children.every((child) => NAMES_NOTHING.includes(child.name));
}

/** A `w:sdtPr` left with nothing inside it is still written, because a control without one is not one we read back */
function emptyProps(props: Props): string {
  return props.attrs ? `<${props.tag} ${props.attrs}/>` : `<${props.tag}/>`;
}

/**
 * The opening of a control with a few children of its `w:sdtPr` swapped out and everything else
 * it carries left exactly as it came.
 * null for a prefix whose shape cannot be made out, which leaves the caller to back out.
 */
export function editSdtPrefix(
  prefix: string,
  edits: readonly SdtPrEdit[]
): string | null {
  const sdt = parseProps(`${prefix}</w:sdt>`);
  const sdtPr = sdt && propsChild(sdt.children, "sdtPr");
  const props = sdtPr ? parseProps(sdtPr.xml) : null;
  if (!sdt || !props) return null;

  const children = edits.reduce(
    (kept, [name, xml]) => setPropsChild(kept, name, xml, SDT_PR_ORDER),
    props.children
  );
  const rendered = renderProps({ ...props, children });
  return renderPrefix({
    ...sdt,
    children: setPropsChild(
      sdt.children,
      "sdtPr",
      rendered === "" ? emptyProps(props) : rendered,
      SDT_ORDER
    ),
  });
}

/**
 * The opening of a control that is going out for the second time, under a name of its own.
 *
 * Two controls sharing an id are merely untidy, but two sharing a `w:dataBinding` are a
 * document Word keeps in step by itself: typing in the one would change the other. So the
 * binding stays with the first copy alone.
 *
 * A prefix whose shape cannot be made out is handed back untouched, since writing the copy
 * as it came is still better than refusing to export at all.
 */
export function copiedControlPrefix(prefix: string, id: number): string {
  const copied = editSdtPrefix(prefix, [
    ["id", `<w:id w:val="${id}"/>`],
    ["dataBinding", null],
  ]);
  return copied ?? prefix;
}
