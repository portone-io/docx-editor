/**
 * Reading the `w:sdt` content controls a document wraps its settled parts in, and naming a
 * control that has to go out more than once.
 *
 * Only the shape we can write back out unchanged is taken apart: the `w:` prefix on both tags
 * we write ourselves, a `w:sdtPr`, and a `w:sdtContent` that carries no attribute of its own
 * and stands last. Everything else is null, which leaves whatever held the control preserved
 * together with its original fragment.
 *
 * What sits inside the control is the caller's business. A whole cell or row
 * (`docx/importTable`), the blocks under a body or a cell (`docx/importSdtBlock`) and a stretch of
 * runs inside a paragraph (`docx/importParagraph`) each read the content their own way, and the
 * wrapper they put back on export is the same string in all of them.
 */

import type { Node as PMNode } from "prosemirror-model";
import { elementXml, openTagXml } from "../ooxml/element";
import { DocxExportError } from "../ooxml/errors";
import { NAMESPACES, wName } from "../ooxml/names";
import {
  type Props,
  parseProps,
  propsChild,
  renderElement,
  setChild,
} from "../ooxml/props";
import { isOnElement, wAttr } from "../ooxml/units";
import { attrString, elementChildren, serializeXml, W_NS } from "../ooxml/xml";
import type { ControlFacts, ControlFlags } from "../schema/controlAttrs";

/**
 * What a control's `w:sdtPr` states, which is all the editor reads off a control besides the XML
 * it goes back out as. The carriers write it under names of their own (`schema/controlAttrs`).
 */
export type SdtFacts = ControlFlags;

/** The opening of a content control taken apart, the content it wraps, and what it says */
export interface SdtWrapper extends SdtFacts {
  /** The `<w:sdt>` opening tag followed by everything that stood ahead of `<w:sdtContent>` */
  prefix: string;
  content: Element;
}

/**
 * The `w:lock` values that shut each of the two clauses a lock settles (§17.18.49). The two are
 * independent, so a control may be un-editable yet removable (`contentLocked`) or editable yet
 * not removable (`sdtLocked`).
 */
const CONTENTS_LOCKED = ["contentLocked", "sdtContentLocked"];
const DELETION_LOCKED = ["sdtLocked", "sdtContentLocked"];

/** The element name with its namespace, which is how a `w:` child is told from a `w14:` one */
function qualifiedName(el: Element): string {
  return `{${el.namespaceURI ?? ""}}${el.localName}`;
}

const GROUP = `{${W_NS}}group`;
const TEMPORARY = `{${W_NS}}temporary`;
const SHOWING_PLACEHOLDER = `{${W_NS}}showingPlcHdr`;

/** Whether this boolean property of the control states on (§17.17.4) */
function states(children: readonly Element[], qualified: string): boolean {
  return isOnElement(
    children.find((child) => qualifiedName(child) === qualified) ?? null
  );
}

/**
 * Everything the editor judges a control by, read out of its properties in one place so that the
 * inline mark, a wrapped cell or row and the block container are all given the same reading.
 *
 * `w:group` (§17.5.2.17) shuts the contents whatever the `w:lock` says, so it is carried beside
 * the two clauses rather than folded into them: the lock is the editor's to lift and the group is
 * not (`spec/notes/contentControls.md`).
 *
 * `w:temporary` and `w:showingPlcHdr` are the two the editor cannot answer by preserving the
 * prefix: the first says the wrapper goes once its contents are edited, the second says what
 * stands inside is placeholder text and is to be shown as such again when the file is opened.
 * Both are read here so that the block, the mark and a wrapped cell or row hear the same thing.
 */
function sdtFacts(sdtPr: Element): SdtFacts {
  const children = elementChildren(sdtPr);
  const lock = children.find((child) => child.localName === "lock");
  const val = lock ? wAttr(lock, "val") : null;
  return {
    contentsLocked: val !== null && CONTENTS_LOCKED.includes(val),
    deletionLocked: val !== null && DELETION_LOCKED.includes(val),
    group: children.some((child) => qualifiedName(child) === GROUP),
    temporary: states(children, TEMPORARY),
    showingPlaceholder: states(children, SHOWING_PLACEHOLDER),
  };
}

/**
 * Whether this child of a `w:sdt` is the content element, which both readers ask the same way so
 * that neither can take a control the other turned down. The prefix a control goes out under is
 * read off whatever stood ahead of this element, so what counts as it decides where that cut falls.
 */
function isSdtContent(child: Element): boolean {
  return child.localName === "sdtContent";
}

/**
 * The opening a control goes back out as: its own tag as it arrived, followed by everything that
 * stood ahead of its content element.
 *
 * This is the exact string the writer splices back into the file (`sdtOpeningXml`), so it is built
 * here alone, for a control whose contents are read as well as for one carrying none.
 */
function sdtPrefix(el: Element, head: readonly Element[]): string {
  return (
    openTagXml(wName("sdt"), attrString(el)) + head.map(serializeXml).join("")
  );
}

/** Takes a `w:sdt` apart into the wrapper to put back on export and the content to read. null for a shape we do not write ourselves */
export function readSdtWrapper(el: Element): SdtWrapper | null {
  if (el.nodeName !== "w:sdt") return null;
  const children = elementChildren(el);
  const contentAt = children.findIndex(isSdtContent);
  const content = contentAt === -1 ? null : children[contentAt];
  // We write the content tag ourselves, so it may carry no other prefix, nothing may hang off it
  // and nothing may follow it
  if (!content || content.nodeName !== "w:sdtContent") return null;
  if (content.attributes.length > 0) return null;
  if (contentAt !== children.length - 1) return null;

  const head = children.slice(0, contentAt);
  const sdtPr = head.find((child) => child.localName === "sdtPr");
  if (!sdtPr) return null;

  return {
    prefix: sdtPrefix(el, head),
    content,
    ...sdtFacts(sdtPr),
  };
}

/**
 * A `w:sdt` that carries no `w:sdtContent` at all, as the facts its carrier records. null for
 * anything else, this editor's own shape of control included.
 *
 * `CT_Sdt` writes the content element `minOccurs="0"`, so a control stating that what it stood
 * around is not there may leave it out entirely, and §17.5.2.34 makes that element a cache of
 * contents rather than the statement itself. `readSdtWrapper` turns such a control down and stays
 * that way: the inline mark and a wrapped cell or row all read the content it hands back, and
 * there is none here for them to stand around. Everything the control carries, `w:sdtEndPr`
 * included, rides in the prefix, exactly as it does when a content element follows it.
 */
export function readEmptySdt(el: Element): ControlFacts | null {
  if (el.nodeName !== "w:sdt") return null;
  const children = elementChildren(el);
  if (children.some(isSdtContent)) return null;
  const sdtPr = children.find((child) => child.localName === "sdtPr");
  if (!sdtPr) return null;
  return {
    prefix: sdtPrefix(el, children),
    ...sdtFacts(sdtPr),
  };
}

/**
 * What the carrier of this control records about it (`schema/controlAttrs`).
 *
 * Every carrier derives it from the wrapper here rather than from the element again, so the inline
 * mark, the block container, the node a control holding nothing stands as, a wrapped cell and a
 * wrapped row cannot come to disagree about what one and the same `w:sdtPr` says.
 */
export function controlFactsFrom(wrapper: SdtWrapper): ControlFacts {
  return {
    prefix: wrapper.prefix,
    contentsLocked: wrapper.contentsLocked,
    deletionLocked: wrapper.deletionLocked,
    group: wrapper.group,
    temporary: wrapper.temporary,
    showingPlaceholder: wrapper.showingPlaceholder,
  };
}

/**
 * What one `w:sdt` holds, which is the question the block reader and the inline reader must
 * answer the same way.
 *
 * A control holding nothing takes two shapes - an empty `w:sdtContent`, and no content element at
 * all, which §17.5.2.34 admits because that element is a cache of what stood inside rather than
 * the statement itself - and the two state the same thing. They are told apart here rather than in
 * each reader, so neither level can come to read one of them as content and the other as nothing.
 */
export type SdtContents =
  | { kind: "nothing"; facts: ControlFacts }
  | { kind: "content"; wrapper: SdtWrapper };

/** What this control holds, and null for a shape this editor does not write back */
export function readSdtContents(el: Element): SdtContents | null {
  const wrapper = readSdtWrapper(el);
  if (!wrapper) {
    const facts = readEmptySdt(el);
    return facts === null ? null : { kind: "nothing", facts };
  }
  return elementChildren(wrapper.content).length === 0
    ? { kind: "nothing", facts: controlFactsFrom(wrapper) }
    : { kind: "content", wrapper };
}

/**
 * The control types whose content the specification restrains to a single run: `w:text`
 * (§17.5.2.44), `w:picture` (§17.5.2.24), `w:date` (§17.5.2.7), `w:comboBox` (§17.5.2.5),
 * `w:dropDownList` (§17.5.2.15), and `w14:checkbox`, which the 2010 extension gives the same
 * shape. A container taking any run of blocks would let a single Enter break that restraint.
 */
const RESTRAINED_TYPES: readonly string[] = [
  `{${W_NS}}text`,
  `{${W_NS}}picture`,
  `{${W_NS}}date`,
  `{${W_NS}}comboBox`,
  `{${W_NS}}dropDownList`,
  `{${NAMESPACES.w14}}checkbox`,
];

/**
 * Whether this control's content may be read as a container of blocks, which is read off the
 * `w:sdtPr` alone so that a control kept whole never has its prefix serialized.
 *
 * Every other type - `w:richText`, which is also what a control naming no type is, `w:group`, the
 * gallery and field types, and the extension types a closed `CT_SdtPr` can only carry as ignorable
 * markup - says what the control means rather than what it may hold.
 */
export function modelsBlockContent(el: Element): boolean {
  const sdtPr = elementChildren(el).find(
    (child) => child.localName === "sdtPr"
  );
  if (!sdtPr) return true;
  return !elementChildren(sdtPr).some((child) =>
    RESTRAINED_TYPES.includes(qualifiedName(child))
  );
}

/**
 * The `w:sdtContent` this editor writes itself, and the control it closes after it.
 *
 * `readSdtWrapper` takes back apart exactly this shape, so the two stand together: a control whose
 * content tag carried an attribute, or did not stand last, is one this never writes.
 */
export const SDT_CLOSING_XML = "</w:sdtContent></w:sdt>";

/** The control's opening tag as it arrived, up to the content tag. A wrapper that lost it cannot go out */
export function sdtOpeningXml(prefix: unknown): string {
  if (typeof prefix !== "string") {
    throw new DocxExportError(
      "lost-original",
      "a content control has lost the opening XML it goes back out as"
    );
  }
  return `${prefix}<w:sdtContent>`;
}

/** One whole control: the opening it arrived with, around content this writer owns */
export function sdtXml(prefix: unknown, content: string): string {
  return sdtOpeningXml(prefix) + content + SDT_CLOSING_XML;
}

/**
 * A control holding nothing, wherever it stood: the opening it arrived with around a content tag
 * with nothing in it.
 *
 * A control that arrived writing no content element at all gains one here: §17.5.2.34 makes that
 * element a cache of what stood inside, so the two shapes state the same thing.
 */
export function serializeEmptyControl(node: PMNode): string {
  return sdtXml(node.attrs.sdtPrefix, "");
}

/**
 * The number Word writes on every control. Nothing reads it back and it only has to differ
 * from the other controls in the document, so a draw out of the whole 32 bit range is enough.
 */
export function newControlId(): number {
  return Math.floor(Math.random() * 0x7fffffff);
}

function renderPrefix(props: Props): string {
  const inner = props.children
    .map((child) => (child.before ?? "") + child.xml)
    .join("");
  return openTagXml(props.tag, props.attrs) + inner + (props.tail ?? "");
}

/** What text one child of the control's `w:sdtPr` is to be changed to. A null xml removes that child */
export type SdtPrEdit = readonly [name: string, xml: string | null];

/** What a control carries when it says nothing but which control it is and whether it is shut */
const NAMES_NOTHING: readonly string[] = ["id", "lock"];

/** A control's opening XML taken apart: the `w:sdt` itself and the `w:sdtPr` it carries */
interface PrefixProps {
  sdt: Props;
  sdtPr: Props;
}

/**
 * The opening XML a control goes back out as, taken apart for reading or for rewriting. null for a
 * shape that cannot be made out, which leaves every caller to back out rather than guess.
 */
function prefixProps(prefix: string): PrefixProps | null {
  const sdt = parseProps(`${prefix}</w:sdt>`);
  const child = sdt && propsChild(sdt.children, "sdtPr");
  const sdtPr = child ? parseProps(child.xml) : null;
  return sdt && sdtPr ? { sdt, sdtPr } : null;
}

/**
 * Whether the control says nothing about itself beyond its id and its lock: no alias or tag, no
 * `w:dataBinding`, no type of its own.
 * A lock may take such a control over a wider stretch, since the stretch is all it ever meant. One
 * Word named keeps the stretch it was given, or the name would come to cover other text.
 */
export function namesNothing(prefix: string): boolean {
  const props = prefixProps(prefix);
  if (!props || props.sdt.attrs !== null || props.sdt.children.length !== 1) {
    return false;
  }
  return props.sdtPr.children.every((child) =>
    NAMES_NOTHING.includes(child.name)
  );
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
  const read = prefixProps(prefix);
  if (!read) return null;
  const { sdt, sdtPr: props } = read;

  // A `w:sdtPr` left with nothing inside it is still written, because a control without one is
  // not one we read back
  const rendered = renderElement(
    edits.reduce((kept, [name, xml]) => setChild(kept, name, xml), props)
  );
  return renderPrefix(setChild(sdt, "sdtPr", rendered));
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
    ["id", elementXml(wName("id"), [[wName("val"), `${id}`]])],
    ["dataBinding", null],
  ]);
  return copied ?? prefix;
}
