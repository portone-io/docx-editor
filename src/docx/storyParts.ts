/**
 * The side stories a part beside the body is written from, told apart by what an edit did to each.
 *
 * Whether a story has to be written again is one question every story writer asks, so it is asked
 * here once. A writer answering it for itself would be free to keep a story another writer drops,
 * and the export invariants would have a third answer of their own.
 */

import type { Node as PMNode } from "prosemirror-model";
import { sameSource } from "../schema/sourceEquality";
import {
  asStoryKey,
  type StoryKey,
  type StoryKind,
  storiesOf,
  storyNodeOf,
} from "../schema/stories";
import type { SessionStore } from "./session";
import type { ImportedStory } from "./story";

export type StoryChange =
  | { readonly change: "kept"; readonly imported: ImportedStory }
  | {
      readonly change: "edited";
      readonly imported: ImportedStory;
      readonly current: PMNode;
    }
  | { readonly change: "removed"; readonly imported: ImportedStory }
  | {
      readonly change: "added";
      readonly key: StoryKey;
      readonly current: PMNode;
    };

const DECIMAL = /^-?\d+$/;

/** Numbers in numeric order, which is the order an entry id is counted in; anything else by its text */
function byId(a: string, b: string): number {
  if (DECIMAL.test(a) && DECIMAL.test(b)) return Number(a) - Number(b);
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

/** The id a key of this kind names */
export function storyIdOf(key: StoryKey, kind: StoryKind): string {
  return key.slice(kind.length + 1);
}

function arrivedChange(doc: PMNode, imported: ImportedStory): StoryChange {
  const current = storyNodeOf(doc, imported.key);
  if (current === null) return { change: "removed", imported };
  return sameSource(current, imported.doc)
    ? { change: "kept", imported }
    : { change: "edited", imported, current };
}

/** Every story of this kind the package arrived with or the document now holds, in part order then id order */
export function storyChangesOf(
  doc: PMNode,
  session: SessionStore,
  kind: StoryKind
): readonly StoryChange[] {
  const arrived = Array.from(session.stories.values())
    .filter((imported) => imported.kind === kind)
    .map((imported) => arrivedChange(doc, imported));
  const added = Object.keys(storiesOf(doc))
    .flatMap((text) => {
      const key = asStoryKey(text);
      return key === null ||
        !key.startsWith(`${kind}:`) ||
        session.stories.has(key)
        ? []
        : [key];
    })
    .sort((a, b) => byId(storyIdOf(a, kind), storyIdOf(b, kind)))
    .flatMap((key): StoryChange[] => {
      const current = storyNodeOf(doc, key);
      return current === null ? [] : [{ change: "added", key, current }];
    });
  return [...arrived, ...added];
}
