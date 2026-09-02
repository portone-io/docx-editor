import {
  DocxEditor,
  type DocxEditorHandle,
  downloadDocx,
} from "@portone/docx-editor";
import { useRef, useState } from "react";
import "@portone/docx-editor/styles.css";
import "./DocxEditorDemo.css";

export interface DocxEditorDemoProps {
  document: File;
}

interface OpenDocument {
  file: File;
  revision: number;
}

export function DocxEditorDemo({ document }: DocxEditorDemoProps) {
  const editorRef = useRef<DocxEditorHandle | null>(null);
  const nextRevision = useRef(0);
  const [opened, setOpened] = useState<OpenDocument>({
    file: document,
    revision: nextRevision.current,
  });
  const [message, setMessage] = useState(document.name);

  const openFile = (file: File) => {
    editorRef.current = null;
    nextRevision.current += 1;
    setOpened({ file, revision: nextRevision.current });
    setMessage(file.name);
  };

  const download = () => {
    try {
      const result = downloadDocx(editorRef.current, {
        fileName: opened.file.name,
      });
      if (result.status === "empty") setMessage("The document is empty.");
      if (result.status === "unavailable") {
        setMessage("The document is still opening.");
      }
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      setMessage(`Could not export the document: ${detail}`);
    }
  };

  return (
    <div className="demo">
      <header className="demo-header">
        <p className="demo-status">Live demo - {message}</p>
        <div className="demo-actions">
          <label className="demo-button">
            Open .docx
            <input
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0];
                if (file) openFile(file);
                event.currentTarget.value = "";
              }}
            />
          </label>
          <button
            className="demo-button-primary"
            type="button"
            onClick={download}
          >
            Export .docx
          </button>
        </div>
      </header>

      <main className="demo-editor">
        <DocxEditor
          key={opened.revision}
          ref={editorRef}
          document={opened.file}
          mode={{
            kind: "edit",
            author: { id: "demo", name: "Demo" },
            locking: true,
          }}
          renderImportError={(error) => (
            <div className="demo-error" role="alert">
              <strong>This document could not be opened.</strong>
              <span>{error.code}</span>
            </div>
          )}
        />
      </main>
    </div>
  );
}
