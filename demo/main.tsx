import { createRoot } from "react-dom/client";
import demoUrl from "../__fixtures__/demo.docx?url";
import { DocxEditorDemo } from "./DocxEditorDemo";
import "./styles.css";

const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

async function loadDemoFile(): Promise<File> {
  const response = await fetch(demoUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const blob = await response.blob();
  return new File([blob], "demo.docx", { type: DOCX_MIME_TYPE });
}

const root = document.getElementById("root");
if (root) {
  const app = createRoot(root);
  app.render(<div className="demo-loading">Opening the demo…</div>);
  void loadDemoFile().then(
    (demoFile) => app.render(<DocxEditorDemo document={demoFile} />),
    (error: unknown) => {
      const detail = error instanceof Error ? error.message : String(error);
      app.render(
        <div className="demo-error" role="alert">
          <strong>Could not load demo.docx.</strong>
          <span>{detail}</span>
        </div>
      );
    }
  );
}
