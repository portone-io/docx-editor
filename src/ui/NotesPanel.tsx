import type { EditorState } from "prosemirror-state";
import type { ReactElement } from "react";
import { documentNotes } from "../editor/commands/noteQueries";
import { editorClassNames } from "../styles/classNames";

export function NotesPanel({
  state,
  pageWidth,
}: {
  state: EditorState;
  pageWidth: number;
}): ReactElement | null {
  const notes = documentNotes(state);
  if (notes.length === 0) return null;

  return (
    <section
      className={editorClassNames.notesPanel}
      aria-label="Document notes"
      style={{ width: `${pageWidth}px` }}
    >
      <h2 className={editorClassNames.notesHeading}>Footnotes and endnotes</h2>
      <ol className={editorClassNames.notesList}>
        {notes.map((note) => (
          <li key={`${note.kind}:${note.id}`}>
            <span className={editorClassNames.noteLabel}>
              {note.kind === "footnote" ? "Footnote" : "Endnote"} {note.label}
            </span>
            <p className={editorClassNames.noteBody}>{note.text}</p>
          </li>
        ))}
      </ol>
    </section>
  );
}
