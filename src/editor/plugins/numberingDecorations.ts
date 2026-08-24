/**
 * Draws list markers and inherited indentation as decorations so display-only values never create
 * new OOXML on export. The plugin also exposes numbering definitions to list commands.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { type ParagraphFormat, toParagraphFormat } from "../../model/format";
import { computeMarkers } from "../../numbering/markers";
import {
  EMPTY_NUMBERING,
  type LevelIndentPt,
  type Numbering,
} from "../../numbering/parseNumbering";
import { editorAttributes, editorCssVariables } from "../../styles/classNames";

/** Word's default (360 twip), used when neither the list definition nor the paragraph has a hanging indent */
const FALLBACK_MARKER_WIDTH_PT = 18;

interface ParagraphSpot {
  pos: number;
  nodeSize: number;
  format: ParagraphFormat | null;
}

function paragraphSpots(doc: PMNode): ParagraphSpot[] {
  const spots: ParagraphSpot[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    spots.push({
      pos,
      nodeSize: node.nodeSize,
      format: toParagraphFormat(node.attrs.format),
    });
    return false;
  });
  return spots;
}

/** A number to draw on screen and the paragraph position it attaches to */
export interface PlacedMarker {
  from: number;
  to: number;
  text: string;
  /** The width of the hanging indent the number sits in */
  widthPt: number;
  /** Logical indentation taken from the level and overlaid only when the paragraph writes none. */
  indentStartPt: number | null;
  indentEndPt: number | null;
  /** The first-line indentation, under the same rule */
  textIndentPt: number | null;
}

/**
 * The indentation of a numbered paragraph and where its number goes.
 *
 * If the paragraph wrote down a `w:ind`, that value wins. It is already applied on screen as
 * paragraph formatting, so we add nothing here (null).
 * Only where nothing was written down do we fill in the value the level specifies, as Word does.
 * The number's slot is exactly the hanging indent width in effect, and Word's default when
 * there is none anywhere.
 */
function markerPlacement(
  format: ParagraphFormat | null,
  level: LevelIndentPt
): Pick<
  PlacedMarker,
  "widthPt" | "indentStartPt" | "indentEndPt" | "textIndentPt"
> {
  const ownTextIndentPt = format?.textIndentPt;
  const textIndentPt = ownTextIndentPt ?? level.textIndentPt;
  const hasOwnLeadingIndent =
    format?.indentLeftPt !== undefined || format?.indentStartPt !== undefined;
  const hasOwnTrailingIndent =
    format?.indentRightPt !== undefined || format?.indentEndPt !== undefined;
  return {
    widthPt:
      textIndentPt !== null && textIndentPt !== undefined && textIndentPt < 0
        ? -textIndentPt
        : FALLBACK_MARKER_WIDTH_PT,
    indentStartPt: hasOwnLeadingIndent ? null : level.startPt,
    indentEndPt: hasOwnTrailingIndent ? null : level.endPt,
    textIndentPt: ownTextIndentPt === undefined ? level.textIndentPt : null,
  };
}

/** Walks a document to find the paragraphs that get numbers and what those numbers are */
export function paragraphMarkers(
  doc: PMNode,
  numbering: Numbering
): PlacedMarker[] {
  const spots = paragraphSpots(doc);
  const markers = computeMarkers(
    spots.map((spot) => spot.format?.numbering ?? null),
    numbering
  );
  return spots.flatMap((spot, index) => {
    const marker = markers[index];
    if (!marker) return [];
    return [
      {
        from: spot.pos,
        to: spot.pos + spot.nodeSize,
        text: marker.text,
        ...markerPlacement(spot.format, marker.indent),
      },
    ];
  });
}

/**
 * The CSS to overlay on a numbered paragraph.
 * A decoration's style is appended after the style the paragraph already carries, so what is
 * written here wins. That is why indentation the paragraph specified itself is not put in here.
 */
function markerStyle(marker: PlacedMarker): string {
  const css = [`${editorCssVariables.markerWidth}:${marker.widthPt}pt`];
  if (marker.indentStartPt !== null) {
    css.push(`margin-inline-start:${marker.indentStartPt}pt`);
  }
  if (marker.indentEndPt !== null) {
    css.push(`margin-inline-end:${marker.indentEndPt}pt`);
  }
  if (marker.textIndentPt !== null) {
    css.push(`text-indent:${marker.textIndentPt}pt`);
  }
  return css.join(";");
}

export function markerDecorations(
  doc: PMNode,
  numbering: Numbering
): DecorationSet {
  const decorations = paragraphMarkers(doc, numbering).map((marker) =>
    Decoration.node(marker.from, marker.to, {
      [editorAttributes.listMarker]: marker.text,
      style: markerStyle(marker),
    })
  );
  return DecorationSet.create(doc, decorations);
}

/** What the plugin holds: the document's list definitions and the numbers to draw right now */
interface NumberingState {
  numbering: Numbering;
  /** Whether the document has a place (numbering.xml) to write the definition of a new list */
  canStartNewList: boolean;
  decorations: DecorationSet;
}

const numberingKey = new PluginKey<NumberingState>("docxEditorNumbering");

/** This document's list definitions. Empty when the editor does not know them */
export function documentNumbering(state: EditorState): Numbering {
  return numberingKey.getState(state)?.numbering ?? EMPTY_NUMBERING;
}

/**
 * Whether a new list can be started in this document.
 * A document without numbering.xml has nowhere to write a new definition, so it would be
 * blocked at export time. In such a document the commands that start a list do not apply
 * in the first place.
 */
export function canStartNewList(state: EditorState): boolean {
  return numberingKey.getState(state)?.canStartNewList ?? true;
}

export function numberingMarkers(
  numbering: Numbering = EMPTY_NUMBERING,
  canStartNewList = true
): Plugin<NumberingState> {
  return new Plugin<NumberingState>({
    key: numberingKey,
    state: {
      init: (_config, state) => ({
        numbering,
        canStartNewList,
        decorations: markerDecorations(state.doc, numbering),
      }),
      apply: (tr, current) =>
        tr.docChanged
          ? {
              numbering,
              canStartNewList,
              decorations: markerDecorations(tr.doc, numbering),
            }
          : current,
    },
    props: {
      decorations: (state) => numberingKey.getState(state)?.decorations,
    },
  });
}
