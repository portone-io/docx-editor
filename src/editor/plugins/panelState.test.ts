import { EditorState, Plugin, PluginKey } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { docxSchema } from "../../schema";
import { panelPlugin } from "./panelState";

interface Spot {
  from: number;
  to: number;
}

function isSpot(value: unknown): value is Spot {
  if (typeof value !== "object" || value === null) return false;
  const { from, to }: Partial<Spot> = value;
  return typeof from === "number" && typeof to === "number";
}

/** A document holding one paragraph of "alphabet", so a spot inside it can be named by number */
function paragraph(): EditorState["doc"] {
  return docxSchema.node("doc", null, [
    docxSchema.node("paragraph", null, [docxSchema.text("alphabet")]),
  ]);
}

/** A flag any transaction can raise, standing in for the protection switch a real panel watches */
const shutKey = new PluginKey<boolean>("panelStateTestShut");

const shutFlag = new Plugin<boolean>({
  key: shutKey,
  state: {
    init: () => false,
    apply: (tr, held) => {
      const meta: unknown = tr.getMeta(shutKey);
      return typeof meta === "boolean" ? meta : held;
    },
  },
});

function stateWith(...plugins: Plugin[]): EditorState {
  return EditorState.create({ doc: paragraph(), plugins });
}

describe("a panel built on the factory", () => {
  it("opens with an anchor and answers it back", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestOpens",
      isAnchor: isSpot,
      onDocChange: "close",
    });
    let state = stateWith(panel.plugin);
    expect(panel.anchor(state)).toBeNull();

    expect(
      panel.open({ from: 1, to: 4 })(state, (tr) => {
        state = state.apply(tr);
      })
    ).toBe(true);
    expect(panel.anchor(state)).toEqual({ from: 1, to: 4 });

    expect(
      panel.close(state, (tr) => {
        state = state.apply(tr);
      })
    ).toBe(true);
    expect(panel.anchor(state)).toBeNull();
    // Nothing left to close, so the press that asked falls through to whatever is behind it
    expect(panel.close(state)).toBe(false);
  });

  it("closes on a document change when told to", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestClosesOnEdit",
      isAnchor: isSpot,
      onDocChange: "close",
    });
    let state = stateWith(panel.plugin);
    panel.open({ from: 1, to: 4 })(state, (tr) => {
      state = state.apply(tr);
    });

    // A selection change is no edit, so the panel stands through it
    state = state.apply(state.tr.setMeta("irrelevant", true));
    expect(panel.anchor(state)).toEqual({ from: 1, to: 4 });

    state = state.apply(state.tr.insertText("x", 1, 1));
    expect(panel.anchor(state)).toBeNull();
  });

  it("carries the anchor through a document change when told to", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestCarries",
      isAnchor: isSpot,
      onDocChange: (anchor, tr) => {
        const from = tr.mapping.map(anchor.from, 1);
        const to = tr.mapping.map(anchor.to, -1);
        return to > from ? { from, to } : null;
      },
    });
    let state = stateWith(panel.plugin);
    // "pha" inside "alphabet"
    panel.open({ from: 3, to: 6 })(state, (tr) => {
      state = state.apply(tr);
    });

    // Two characters typed ahead of it push the whole anchor along
    state = state.apply(state.tr.insertText("xy", 1, 1));
    expect(panel.anchor(state)).toEqual({ from: 5, to: 8 });

    // And the edit that swallows the anchored text closes the panel
    state = state.apply(state.tr.delete(5, 8));
    expect(panel.anchor(state)).toBeNull();
  });

  it("open reports false where canOpen says no", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestCanOpen",
      isAnchor: isSpot,
      onDocChange: "close",
      canOpen: (_state, anchor) => anchor.to > anchor.from,
    });
    let state = stateWith(panel.plugin);
    let dispatched = 0;
    const dispatch = (tr: EditorState["tr"]) => {
      dispatched += 1;
      state = state.apply(tr);
    };

    expect(panel.open({ from: 2, to: 2 })(state, dispatch)).toBe(false);
    expect(dispatched).toBe(0);
    expect(panel.anchor(state)).toBeNull();

    expect(panel.open({ from: 2, to: 5 })(state, dispatch)).toBe(true);
    expect(panel.anchor(state)).toEqual({ from: 2, to: 5 });
  });

  it("closes on a meta-only transaction where closeWhen says so", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestCloseWhen",
      isAnchor: isSpot,
      onDocChange: "close",
      closeWhen: (state) => shutKey.getState(state) === true,
    });
    // The flag leads the panel, so the panel reads the value this very transaction set
    let state = stateWith(shutFlag, panel.plugin);
    panel.open({ from: 1, to: 4 })(state, (tr) => {
      state = state.apply(tr);
    });
    expect(panel.anchor(state)).toEqual({ from: 1, to: 4 });

    const shutting = state.tr.setMeta(shutKey, true);
    expect(shutting.docChanged).toBe(false);
    state = state.apply(shutting);
    expect(panel.anchor(state)).toBeNull();
  });

  it("answers nothing for a state built without it, and opens over nothing it cannot read back", () => {
    const panel = panelPlugin<Spot>({
      name: "panelStateTestAbsent",
      isAnchor: isSpot,
      onDocChange: "close",
    });
    expect(panel.anchor(stateWith())).toBeNull();
    expect(panel.open({ from: 1, to: 4 })(stateWith())).toBe(true);

    // Metadata under the panel's own key that is no anchor of its own leaves it shut
    let state = stateWith(panel.plugin);
    state = state.apply(state.tr.setMeta(panel.plugin, { from: 1 }));
    expect(panel.anchor(state)).toBeNull();

    panel.open({ from: 1, to: 4 })(state, (tr) => {
      state = state.apply(tr);
    });
    state = state.apply(state.tr.setMeta(panel.plugin, "open please"));
    expect(panel.anchor(state)).toEqual({ from: 1, to: 4 });
  });
});
