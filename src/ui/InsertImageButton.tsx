/** Toolbar button backed by a hidden native image file input. */

import { Image } from "lucide-react";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { type ReactElement, useRef } from "react";
import {
  IMAGE_FILE_ACCEPT,
  imageFilesIn,
  insertImageFiles,
} from "../editor/imageFiles";
import { canInsertImage } from "../editor/insertImage";
import { ToolbarButton } from "./ToolbarButton";

export interface InsertImageButtonProps {
  view: EditorView;
  state: EditorState;
}

export function InsertImageButton({
  view,
  state,
}: InsertImageButtonProps): ReactElement {
  const pickerRef = useRef<HTMLInputElement | null>(null);

  return (
    <>
      <ToolbarButton
        label="Insert image"
        icon={Image}
        disabled={!view.editable || !canInsertImage(state)}
        onRun={() => pickerRef.current?.click()}
      />
      <input
        ref={pickerRef}
        type="file"
        accept={IMAGE_FILE_ACCEPT}
        hidden
        onChange={async (event) => {
          const picker = event.currentTarget;
          const files = imageFilesIn(picker);
          // Picking the same file twice in a row has to raise a change event the second time too
          picker.value = "";
          await insertImageFiles(view, files);
          // Typing has to be able to continue where the image landed
          if (!view.isDestroyed) view.focus();
        }}
      />
    </>
  );
}
