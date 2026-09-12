/**
 * How a menu row spells the modifiers of a key binding.
 *
 * `Mod` is Command on Apple platforms and Control elsewhere, and which one a binding answers to is
 * decided by `prosemirror-keymap` off the platform name. A label is read off the same platform
 * test, so a row cannot promise a key the bindings do not hold - an iPad keyboard sends Command
 * for every `Mod` binding the editor declares (`editor/clipboard/plugin` reads it for the same
 * reason).
 */

const APPLE = /Mac|iP(hone|[oa]d)/;

export interface ModifierLabels {
  /** `Mod`, as a row writes it before the key */
  readonly mod: string;
  readonly alt: string;
}

/** The spellings this platform's bindings are labelled with */
export function modifierLabels(): ModifierLabels {
  const platform = typeof navigator === "undefined" ? "" : navigator.platform;
  return APPLE.test(platform)
    ? { mod: "⌘", alt: "⌥" }
    : { mod: "Ctrl+", alt: "Alt+" };
}
