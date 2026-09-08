import type { Node as PMNode } from "prosemirror-model";

type MarkerAttr = "leadingXml" | "trailingXml";

/**
 * Whether this node is content the editor keeps as it came rather than one it models.
 *
 * The block placeholder declares it as a group; the two inline ones cannot, since their group says
 * where they may stand, so they are named here and asked for by name everywhere else.
 */
export function isPreservedNode(node: PMNode): boolean {
  return (
    node.type.name === "rawInline" ||
    node.type.name === "rawRunContent" ||
    node.type.isInGroup("preserved")
  );
}

/** Visit preserved XML in the order the writer places it around editable content. */
export function visitPreservedFragments(
  doc: PMNode,
  visit: (
    node: PMNode,
    pos: number,
    xml: string | null,
    attr?: MarkerAttr
  ) => void
): void {
  const visitAttr = (node: PMNode, pos: number, attr: MarkerAttr) => {
    const xml = node.attrs[attr];
    if (typeof xml === "string" && xml.length > 0) {
      visit(node, pos, xml, attr);
    }
  };
  const walk = (node: PMNode, pos: number) => {
    if (isPreservedNode(node)) {
      visit(
        node,
        pos,
        typeof node.attrs.xml === "string" ? node.attrs.xml : null
      );
    }
    if (node.type.name === "table" || node.type.name === "tableRow") {
      visitAttr(node, pos, "leadingXml");
    }
    node.forEach((child, offset) => {
      walk(child, pos + 1 + offset);
    });
    if (node.type.name === "tableRow" || node.type.name === "tableCell") {
      visitAttr(node, pos, "trailingXml");
    }
  };
  doc.descendants((node, pos) => {
    walk(node, pos);
    return false;
  });
}
