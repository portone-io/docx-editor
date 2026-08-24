/**
 * The command that puts an image into the document.
 *
 * What it takes is what an image node carries: the bytes as a data URL and the size the
 * drawing records, in EMU. Working out that size from the image itself takes a decoded
 * image, which only a browser can give, so it happens where the file is read
 * (`editor/imageFiles.ts`) and never in here.
 *
 * The node is created without original XML, and that is what marks it as inserted: the
 * export writes a media part for its bytes and builds a drawing around them.
 */

import type { Command, EditorState } from "prosemirror-state";
import { insertPoint } from "prosemirror-transform";
import { type ImageExtent, toImageExtent, toImageSrc } from "../ooxml/image";
import { docxSchema } from "../schema";
import { replacementShut } from "../schema/locks";

/** What has to be known about an image before it can go into a document */
export interface ImageToInsert {
  /** The image bytes as a data URL */
  src: string;
  /** The size to show it at, in EMU */
  extent: ImageExtent;
  /** The alternative text. Nothing when there is none */
  alt?: string | null;
}

/**
 * Where the image goes.
 *
 * The selection itself when an inline node can stand there, and the nearest position that
 * can hold one otherwise. Null when no position in this document can, which is the case
 * for a document made up entirely of preserved blocks.
 * Null as well where a lock shuts the selection: the image goes in in place of whatever is
 * selected, so the question is whether the guard would let that replacement through
 * (`schema/locks`), which a control the document locked against deletion answers no to even where
 * its contents stand open.
 */
function insertPosition(state: EditorState): number | null {
  if (replacementShut(state.selection, state.doc)) return null;
  return insertPoint(state.doc, state.selection.from, docxSchema.nodes.image);
}

/** Whether an image can be inserted at the current position. Used to enable and disable the toolbar button */
export function canInsertImage(state: EditorState): boolean {
  return insertPosition(state) !== null;
}

/**
 * Puts the image where the selection is, in place of whatever was selected.
 *
 * It goes in carrying the formatting of the text around it, so the run it ends up in is
 * the one it was inserted into - the same as a picture pasted in Word.
 */
export function insertImage(image: ImageToInsert): Command {
  return (state, dispatch) => {
    const src = toImageSrc(image.src);
    const extent = toImageExtent(image.extent);
    if (src === null || extent === null || insertPosition(state) === null) {
      return false;
    }
    if (dispatch) {
      const node = docxSchema.nodes.image.create({
        src,
        extent,
        alt: image.alt ?? null,
        xml: null,
      });
      dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    }
    return true;
  };
}
