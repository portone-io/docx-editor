/**
 * Draws list markers and inherited indentation as decorations so display-only values never create
 * new OOXML on export. The definitions the markers are drawn from belong to the document
 * snapshot, and the readers list commands ask stand here beside the drawing.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import {
  type ParagraphFormat,
  type RunFormat,
  toParagraphFormat,
} from "../../model/format";
import { computeMarkers } from "../../numbering/markers";
import type {
  LevelAlign,
  LevelIndentPt,
  LevelSuffix,
  Numbering,
} from "../../numbering/parseNumbering";
import { editorAttributes, editorCssVariables } from "../../styles/classNames";
import { documentOf } from "../editorDocument";

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
  /** What the level puts between the number and the text (`w:suff`) */
  suffix: LevelSuffix;
  /** Where the number sits inside the width kept for it (`w:lvlJc`) */
  align: LevelAlign;
  /** The formatting the level puts on the number alone (`lvl/rPr`) */
  run: RunFormat | null;
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
        suffix: marker.suffix,
        align: marker.align,
        run: marker.run,
        ...markerPlacement(spot.format, marker.indent),
      },
    ];
  });
}

/**
 * How much room the number takes before the text begins.
 *
 * A tab suffix is the one that keeps a width for the number, the hanging indent wide, so the text
 * of every item starts at the same place. The other two put the text straight after the number,
 * with the space among them drawn as part of the marker itself.
 */
function markerSpacing(marker: PlacedMarker): string[] {
  if (marker.suffix === "tab") {
    return [`${editorCssVariables.markerWidth}:${marker.widthPt}pt`];
  }
  return [
    `${editorCssVariables.markerWidth}:0`,
    `${editorCssVariables.markerGap}:0`,
  ];
}

/**
 * The character formatting the level puts on its number (§17.9.24).
 *
 * It dresses the number and nothing else, so it goes out as variables the rule drawing the number
 * reads rather than as CSS the paragraph itself wears. An underline, a strikethrough or a
 * highlight would run across the text the number stands in front of, so they are left undrawn.
 */
function markerFormat(run: RunFormat | null): string[] {
  if (!run) return [];
  const css: string[] = [];
  if (run.bold !== undefined) {
    css.push(
      `${editorCssVariables.markerFontWeight}:${run.bold ? "bold" : "normal"}`
    );
  }
  if (run.italic !== undefined) {
    css.push(
      `${editorCssVariables.markerFontStyle}:${run.italic ? "italic" : "normal"}`
    );
  }
  if (run.color !== undefined) {
    css.push(`${editorCssVariables.markerColor}:${run.color}`);
  }
  if (run.fontSizePt !== undefined) {
    css.push(`${editorCssVariables.markerFontSize}:${run.fontSizePt}pt`);
  }
  if (run.fontFamily !== undefined) {
    css.push(`${editorCssVariables.markerFontFamily}:${run.fontFamily}`);
  }
  return css;
}

/**
 * The CSS to overlay on a numbered paragraph.
 * A decoration's style is appended after the style the paragraph already carries, so what is
 * written here wins. That is why indentation the paragraph specified itself is not put in here.
 */
function markerStyle(marker: PlacedMarker): string {
  const css = [...markerSpacing(marker), ...markerFormat(marker.run)];
  if (marker.align !== "left") {
    css.push(`${editorCssVariables.markerAlign}:${marker.align}`);
  }
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

const numberingKey = new PluginKey<DecorationSet>("docxEditorNumbering");

/** This document's list definitions. Empty when the editor does not know them */
export function documentNumbering(state: EditorState): Numbering {
  return documentOf(state).formatting.numbering;
}

/**
 * Whether a new list can be started in this document.
 * A document without numbering.xml has nowhere to write a new definition, so it would be
 * blocked at export time. In such a document the commands that start a list do not apply
 * in the first place.
 */
export function canStartNewList(state: EditorState): boolean {
  return documentOf(state).canStartNewList;
}

/** Draws the numbers this document's lists put in front of their paragraphs */
export function numberingMarkers(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: numberingKey,
    state: {
      init: (_config, state) =>
        markerDecorations(state.doc, documentOf(state).formatting.numbering),
      apply: (tr, current, _old, state) =>
        tr.docChanged
          ? markerDecorations(tr.doc, documentOf(state).formatting.numbering)
          : current,
    },
    props: {
      decorations: (state) => numberingKey.getState(state),
    },
  });
}
