import { useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import demoUrl from "../__fixtures__/demo.docx?url";
import { DocxEditor, type DocxEditorHandle, downloadDocx } from "../src/index";
import "../src/styles/editor.css";
import "./styles.css";

interface OpenDocument {
  file: File;
  revision: number;
}

const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function Playground({ demoFile }: { demoFile: File }) {
  const editorRef = useRef<DocxEditorHandle | null>(null);
  const nextRevision = useRef(0);
  const [opened, setOpened] = useState<OpenDocument>({
    file: demoFile,
    revision: nextRevision.current,
  });
  const [message, setMessage] = useState(demoFile.name);

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
    <div className="playground">
      <header className="playground-header">
        <div>
          <h1>docx-editor</h1>
          <p>{message}</p>
        </div>
        <div className="playground-actions">
          <label className="playground-button">
            Open DOCX
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
          <button type="button" onClick={() => openFile(demoFile)}>
            Reset demo
          </button>
          <button type="button" onClick={download}>
            Download
          </button>
        </div>
      </header>

      <main className="playground-editor">
        <DocxEditor
          key={opened.revision}
          ref={editorRef}
          document={opened.file}
          mode={{ kind: "edit", locking: true }}
          renderImportError={(error) => (
            <div className="playground-error" role="alert">
              <strong>This document could not be opened.</strong>
              <span>{error.code}</span>
            </div>
          )}
        />
      </main>
    </div>
  );
}

async function loadDemoFile(): Promise<File> {
  const response = await fetch(demoUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const blob = await response.blob();
  return new File([blob], "demo.docx", { type: DOCX_MIME_TYPE });
}

const root = document.getElementById("root");
if (root) {
  const app = createRoot(root);
  app.render(<div className="playground-loading">Opening the demo…</div>);
  void loadDemoFile().then(
    (demoFile) => app.render(<Playground demoFile={demoFile} />),
    (error: unknown) => {
      const detail = error instanceof Error ? error.message : String(error);
      app.render(
        <div className="playground-error" role="alert">
          <strong>Could not load demo.docx.</strong>
          <span>{detail}</span>
        </div>
      );
    }
  );
}
