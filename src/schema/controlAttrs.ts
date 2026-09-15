/**
 * The attributes each carrier of a content control writes what the control states in.
 *
 * Three carriers state the same facts: the `sdt` mark and the `sdtBlock` container under the
 * control's own names, and a cell the file wrapped (`docx/importTable`) under `sdt`-prefixed ones,
 * beside the cell's own attributes. The names stand here once, so that a fact the editor reads off
 * a control is declared, drawn, parsed, imported, merged and cleared from one table rather than
 * from a list written out afresh in every module that touches one.
 *
 * Nothing is imported here: the schema itself reads this table (`./docxSchema`), and so does the
 * export, which must reach no editor code at all (`core.test`).
 */

import type { Node as PMNode } from "prosemirror-model";

/** The node a block-level content control is read as, which `./docxSchema` declares it under */
export const SDT_BLOCK_NODE = "sdtBlock";

/** Which attribute one carrier writes each fact under */
export interface ControlAttrNames {
  prefix: string;
  contentsLocked: string;
  deletionLocked: string;
  group: string;
  temporary: string;
  showingPlaceholder: string;
}

/** What a control states, in the shape every carrier reads it as (`docx/sdt`) */
export interface ControlFacts {
  /** The `<w:sdt>` opening tag with everything that stood ahead of `<w:sdtContent>` */
  prefix: string | null;
  /** Whether its `w:lock` says the contents may not be edited */
  contentsLocked: boolean;
  /** Whether its `w:lock` says it may not be deleted, not even whole */
  deletionLocked: boolean;
  /** Whether it is a `w:group` */
  group: boolean;
  /** Whether `w:temporary` says the control goes once its contents are edited (§17.5.2.43) */
  temporary: boolean;
  /** Whether `w:showingPlcHdr` says what stands inside is placeholder text (§17.5.2.39) */
  showingPlaceholder: boolean;
}

/** Everything a control states other than the opening XML the carrier preserves */
export type ControlFlags = Omit<ControlFacts, "prefix">;

/** The control's own names, which the mark and the block-level container both carry */
export const OWN_CONTROL_ATTRS: ControlAttrNames = {
  prefix: "sdtPrefix",
  contentsLocked: "contentsLocked",
  deletionLocked: "deletionLocked",
  group: "group",
  temporary: "temporary",
  showingPlaceholder: "showingPlaceholder",
};

/** A wrapped cell carries the control beside its own attributes, so every name is prefixed */
export const CELL_CONTROL_ATTRS: ControlAttrNames = {
  prefix: "sdtPrefix",
  contentsLocked: "sdtContentsLocked",
  deletionLocked: "sdtDeletionLocked",
  group: "sdtGroup",
  temporary: "sdtTemporary",
  showingPlaceholder: "sdtShowingPlaceholder",
};

/** What a carrier no control stands around says, which is what each of them defaults to */
export const NO_CONTROL: ControlFacts = {
  prefix: null,
  contentsLocked: false,
  deletionLocked: false,
  group: false,
  temporary: false,
  showingPlaceholder: false,
};

/** The DOM attribute each flag is drawn under, which the three carriers share (`./docxSchema`) */
const FLAG_DOM_ATTRS: Readonly<Record<keyof ControlFlags, string>> = {
  contentsLocked: "data-sdt-contents-locked",
  deletionLocked: "data-sdt-deletion-locked",
  group: "data-sdt-group",
  temporary: "data-sdt-temporary",
  showingPlaceholder: "data-sdt-placeholder",
};

/** These facts as the attributes of one carrier, to be spread over its own */
export function controlAttrs(
  names: ControlAttrNames,
  facts: ControlFacts
): Record<string, string | boolean | null> {
  return {
    [names.prefix]: facts.prefix,
    [names.contentsLocked]: facts.contentsLocked,
    [names.deletionLocked]: facts.deletionLocked,
    [names.group]: facts.group,
    [names.temporary]: facts.temporary,
    [names.showingPlaceholder]: facts.showingPlaceholder,
  };
}

/** What one carrier's attributes say the control standing there states */
export function controlFactsOf(
  names: ControlAttrNames,
  attrs: Record<string, unknown>
): ControlFacts {
  const prefix: unknown = attrs[names.prefix];
  return {
    prefix: typeof prefix === "string" ? prefix : null,
    contentsLocked: attrs[names.contentsLocked] === true,
    deletionLocked: attrs[names.deletionLocked] === true,
    group: attrs[names.group] === true,
    temporary: attrs[names.temporary] === true,
    showingPlaceholder: attrs[names.showingPlaceholder] === true,
  };
}

/** What one carrier declares for them, each defaulting to no control standing there */
export function controlAttrSpecs(
  names: ControlAttrNames
): Record<string, { default: string | boolean | null }> {
  return Object.fromEntries(
    Object.entries(controlAttrs(names, NO_CONTROL)).map(([attr, value]) => [
      attr,
      { default: value },
    ])
  );
}

const drawn = (flag: boolean): string | undefined => (flag ? "1" : undefined);

/** The flags as a carrier draws them, its opening XML being drawn beside them */
export function controlFlagsToDom(
  facts: ControlFlags
): Record<string, string | undefined> {
  return {
    [FLAG_DOM_ATTRS.contentsLocked]: drawn(facts.contentsLocked),
    [FLAG_DOM_ATTRS.deletionLocked]: drawn(facts.deletionLocked),
    [FLAG_DOM_ATTRS.group]: drawn(facts.group),
    [FLAG_DOM_ATTRS.temporary]: drawn(facts.temporary),
    [FLAG_DOM_ATTRS.showingPlaceholder]: drawn(facts.showingPlaceholder),
  };
}

/** The flags read back off a carrier that was drawn */
export function controlFlagsFromDom(dom: HTMLElement): ControlFlags {
  const set = (attr: string): boolean => dom.getAttribute(attr) === "1";
  return {
    contentsLocked: set(FLAG_DOM_ATTRS.contentsLocked),
    deletionLocked: set(FLAG_DOM_ATTRS.deletionLocked),
    group: set(FLAG_DOM_ATTRS.group),
    temporary: set(FLAG_DOM_ATTRS.temporary),
    showingPlaceholder: set(FLAG_DOM_ATTRS.showingPlaceholder),
  };
}

/**
 * Whether this node is a block-level content control (`docx/importSdtBlock`).
 *
 * The comparison is by type name, so a schema built beside the editor's own
 * (`table/__testing__/tables`) is read as well.
 */
export function isBlockControl(node: PMNode | null | undefined): boolean {
  return node?.type.name === SDT_BLOCK_NODE;
}

/**
 * Which attributes the control standing at this node writes what it states in, and null where no
 * control stands there.
 *
 * These are the two containers a control stands around whole: a cell the file wrapped
 * (`docx/importTable`), which carries the control's opening XML beside its own attributes, and a
 * block-level control, which is the wrapper itself. A cell no control wrapped carries none of
 * this, and counting it as a control would end the walk out of the tree at the first cell and open
 * every lock standing around the table (`./locks`).
 * A cell is found by its table role, so a schema built beside the editor's own is read as well.
 */
export function controlAttrsOf(
  node: PMNode | null | undefined
): ControlAttrNames | null {
  if (!node) return null;
  if (isBlockControl(node)) return OWN_CONTROL_ATTRS;
  if (node.type.spec.tableRole !== "cell") return null;
  return typeof node.attrs[CELL_CONTROL_ATTRS.prefix] === "string"
    ? CELL_CONTROL_ATTRS
    : null;
}
