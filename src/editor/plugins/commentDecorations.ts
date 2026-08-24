/** Draws the text ranges identified by Word comment markers. */

import { Plugin } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { editorClassNames } from "../../styles/classNames";
import { documentComments } from "../commands/commentCommands";

export function commentDecorations(): Plugin {
  return new Plugin({
    props: {
      decorations(state) {
        const decorations = documentComments(state)
          .filter((comment) => comment.from < comment.to)
          .map((comment) =>
            Decoration.inline(comment.from, comment.to, {
              class: editorClassNames.commentRange,
              "data-comment-id": comment.id,
            })
          );
        return DecorationSet.create(state.doc, decorations);
      },
    },
  });
}
