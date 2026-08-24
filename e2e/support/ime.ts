/**
 * Driving the browser's own IME.
 *
 * `Input.imeSetComposition` is what a kana-kanji or a pinyin conversion does while it is still
 * open, and `Input.insertText` is the commit that follows it. They are CDP methods rather than
 * anything a Playwright keyboard call can reach, because a composition is renderer state and not
 * a sequence of key events.
 */

import type { CDPSession, Page } from "@playwright/test";

export function imeSession(page: Page): Promise<CDPSession> {
  return page.context().newCDPSession(page);
}

/** Sets what the IME is composing right now, with the caret at the end of it */
export function setComposition(
  cdp: CDPSession,
  text: string
): Promise<unknown> {
  return cdp.send("Input.imeSetComposition", {
    text,
    selectionStart: text.length,
    selectionEnd: text.length,
  });
}

/**
 * Empties the open composition, which is what the backspace taking the last jamo of a hangul buffer
 * off it does. An empty buffer is answered with one more `compositionupdate` and then a
 * `compositionend`, so a composition backspaced away is over rather than left standing open.
 */
export function clearComposition(cdp: CDPSession): Promise<unknown> {
  return setComposition(cdp, "");
}

/** Commits the open composition, the way picking a candidate does */
export function commitComposition(
  cdp: CDPSession,
  text: string
): Promise<unknown> {
  return cdp.send("Input.insertText", { text });
}

/** Walks a composition through its stages, one `imeSetComposition` per stage */
export async function compose(
  cdp: CDPSession,
  stages: readonly string[]
): Promise<void> {
  for (const stage of stages) await setComposition(cdp, stage);
}

interface Key {
  key: string;
  code: string;
  keyCode: number;
}

export const ENTER: Key = { key: "Enter", code: "Enter", keyCode: 13 };

/** A full press of one key, the pair of events a keyboard sends */
export async function pressKey(cdp: CDPSession, of: Key): Promise<void> {
  const shared = {
    key: of.key,
    code: of.code,
    windowsVirtualKeyCode: of.keyCode,
    nativeVirtualKeyCode: of.keyCode,
  };
  await cdp.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...shared });
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", ...shared });
}
