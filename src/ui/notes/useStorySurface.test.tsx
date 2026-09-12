// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeNotesDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { openFootnote } from "../../editor/commands/noteCommands";
import { noteProjection } from "../../editor/commands/noteQueries";
import {
  createEditorView,
  editorStateForSession,
} from "../../editor/createEditor";
import { sectionGeometryAt } from "../../editor/documentStyles";
import { storyDocument } from "../../editor/editorDocument";
import { noteExtensions, noteHost } from "../../editor/notes/noteSurface";
import { requestedNote } from "../../editor/plugins/noteNavigation";
import { docxSchema } from "../../schema";
import { storyKey } from "../../schema/stories";
import {
  type StorySurface,
  type StorySurfaceBinding,
  useStorySurface,
} from "./useStorySurface";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const FOOTNOTE = storyKey("footnote", "2");

/** The footnote binding, as the component that mounts the editor declares one */
const FOOTNOTES: StorySurfaceBinding = {
  hostOf: noteHost,
  extensionsOf: (main, row) => noteExtensions(main, row.key, () => row.label),
  documentFor: (main, snapshot, row) =>
    storyDocument(snapshot, sectionGeometryAt(main, row.referencePos)),
  requestedIn: requestedNote,
  openIn: () => {},
};

let host: HTMLDivElement;
let root: Root | null = null;
let main: EditorView | null = null;
let surface: StorySurface | null = null;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  if (root) act(() => root?.unmount());
  root = null;
  main?.destroy();
  main = null;
  surface = null;
  host.remove();
});

function Probe({ state, view }: { state: EditorState; view: EditorView }) {
  surface = useStorySurface({
    main: { view, state },
    rows: noteProjection.read(state).footnotes,
    binding: FOOTNOTES,
  });
  return null;
}

function held(): StorySurface {
  if (surface === null) throw new Error("the hook answered nothing");
  return surface;
}

/** Draws the surface over the state the document is now in, as a state change redraws it */
function show(view: EditorView): void {
  const live = root ?? createRoot(host);
  root = live;
  act(() => live.render(<Probe state={view.state} view={view} />));
}

function openMain(): EditorView {
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(importDocx(makeNotesDocx())),
    onStateChange: () => {},
  });
  main = view;
  return view;
}

/** Where the body's one footnote reference stands */
function referenceAt(view: EditorView): number {
  let found = -1;
  view.state.doc.descendants((node, pos) => {
    if (node.type === docxSchema.nodes.noteReference && found === -1) {
      found = pos;
    }
    return true;
  });
  if (found === -1) throw new Error("the body refers to no note");
  return found;
}

describe("the story surface", () => {
  it("opens the story the document asks for and closes it when its row goes", () => {
    const view = openMain();
    show(view);
    expect(held().open).toBeNull();
    expect(held().editing).toBeNull();
    expect(held().active?.surface).toBe("body");

    act(() => {
      openFootnote("2")(view.state, (tr) => view.dispatch(tr));
    });
    show(view);

    expect(held().open).toBe(FOOTNOTE);
    expect(held().editing).not.toBeNull();
    expect(held().readOnly).toBe(false);

    // The reference is what the footnote hangs from, so deleting it leaves the story nowhere to
    // stand and the caret back in the body
    const at = referenceAt(view);
    act(() => {
      view.dispatch(view.state.tr.delete(at, at + 1));
    });
    show(view);

    expect(held().open).toBeNull();
    expect(held().editing).toBeNull();
    expect(held().active?.surface).toBe("body");
  });
});
