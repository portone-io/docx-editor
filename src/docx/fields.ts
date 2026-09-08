/**
 * Where the fields of a paragraph stand, read back off the fragments the editor keeps as they came.
 *
 * A complex field is not one element. It is a begin character, the instruction text, an optional
 * separate character, the result its producer cached, and an end character, all standing loose
 * among the runs of the paragraph in that order (§17.16.18). Nothing in the model pairs them up, so
 * anything that wants to know what a field was told to print has to read the characters back. A
 * simple field says the same in one element and carries its instruction as an attribute (§17.16.19).
 *
 * What a field means, and how to work it out again, belongs to the fields feature. This answers the
 * one question a projection has: where each field stands and what its instruction says.
 */

import type { Node as PMNode } from "prosemirror-model";
import { parsePropsXml } from "../ooxml/props";
import { attributeByLocalName } from "../ooxml/xml";

export interface FieldSpan {
  /** The instruction text between begin and separate (or end), whitespace-trimmed */
  instr: string;
  /** Where the begin character stands, and the simple field itself for one written as a single element */
  begin: number;
  /** Where the separate character stands, and null for a field that cached no result */
  separate: number | null;
  /** Where the end character stands, which is `begin` again for a simple field */
  end: number;
  /** The half-open range of positions the cached result covers, and null for a field carrying none */
  result: [number, number] | null;
}

/** §17.16.18 `w:fldChar`, the three characters a complex field is written between */
const FIELD_CHARACTER = "fldChar";

/** §17.16.23 `w:instrText`, one piece of what the field was told to print */
const INSTRUCTION = "instrText";

/** §17.16.19 `w:fldSimple`, a whole field written as one element */
const SIMPLE_FIELD = "fldSimple";

/**
 * One attribute of the element a preserved fragment stands for.
 *
 * The fragment holds the element as text rather than as attrs of its own, since the editor has no
 * model for it, so the text is what an attribute is read back out of.
 */
function attributeOf(node: PMNode, name: string): string | null {
  const xml: unknown = node.attrs.xml;
  if (typeof xml !== "string") return null;
  const el = parsePropsXml(xml);
  return el === null ? null : attributeByLocalName(el, name);
}

function chipText(node: PMNode): string {
  const text: unknown = node.attrs.text;
  return typeof text === "string" ? text : "";
}

/** §17.16.24 `w:delInstrText`, an instruction a tracked change took away */
const DELETED_INSTRUCTION = "delInstrText";

const FIELD_PLUMBING: ReadonlySet<string> = new Set([
  FIELD_CHARACTER,
  INSTRUCTION,
  DELETED_INSTRUCTION,
]);

/**
 * Whether the node is one of the pieces a complex field is written out of rather than something
 * the field prints. A projection drops these whether or not they pair up into a field.
 */
export function isFieldCharacter(node: PMNode): boolean {
  const element: unknown = node.attrs.element;
  return (
    node.type.name === "rawRunContent" &&
    typeof element === "string" &&
    FIELD_PLUMBING.has(element)
  );
}

/** A complex field being read, from its begin character up to the end character that closes it */
interface OpenField {
  instr: string;
  begin: number;
  separate: number | null;
}

/**
 * The fields this paragraph holds, in the order they begin, each position offset by `offset`.
 *
 * Positions are counted inside the paragraph's content, so `offset` is where that content starts:
 * a caller reading a paragraph standing on its own passes 0, and one reading a paragraph of a
 * document passes the position just inside it.
 *
 * A begin character with no end is no field: the file says a field starts there and never says
 * what it prints, so there is nothing to pair it with and it is passed over. `documentFidelity`
 * still reports the characters themselves as content kept rather than modelled.
 */
export function fieldSpans(paragraph: PMNode, offset: number): FieldSpan[] {
  const spans: FieldSpan[] = [];
  const open: OpenField[] = [];
  paragraph.forEach((child, at) => {
    const pos = offset + at;
    const element: unknown = child.attrs.element;
    if (child.type.name === "rawInline" && element === SIMPLE_FIELD) {
      spans.push({
        instr: (attributeOf(child, "instr") ?? "").trim(),
        begin: pos,
        separate: null,
        end: pos,
        result: null,
      });
      return;
    }
    if (child.type.name !== "rawRunContent") return;
    if (element === INSTRUCTION) {
      // What follows the separate character is the cached result, not more of the instruction
      const active = open.at(-1);
      if (active && active.separate === null) active.instr += chipText(child);
      return;
    }
    if (element !== FIELD_CHARACTER) return;
    const kind = attributeOf(child, "fldCharType");
    if (kind === "begin") {
      open.push({ instr: "", begin: pos, separate: null });
      return;
    }
    const active = open.at(-1);
    if (active === undefined) return;
    if (kind === "separate") {
      // A second separate says nothing the first did not; the result still starts at the first
      if (active.separate === null) active.separate = pos;
      return;
    }
    if (kind !== "end") return;
    open.pop();
    spans.push({
      instr: active.instr.trim(),
      begin: active.begin,
      separate: active.separate,
      end: pos,
      result: active.separate === null ? null : [active.separate + 1, pos],
    });
  });
  // A nested field ends before the field around it, so the ends are not the order they begin in
  return spans.sort((a, b) => a.begin - b.begin);
}
