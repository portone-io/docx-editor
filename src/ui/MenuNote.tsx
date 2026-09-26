/**
 * A line at the head of a right-click menu saying why its editing entries stand disabled.
 *
 * A menu holds only its items, so the line is presentational and the menu names it as its
 * description instead, which is what a screen reader reads out when the menu opens.
 */

import type { ReactElement } from "react";
import { editorClassNames } from "../styles/classNames";

export const LOCKED_NOTE = "Locked content can't be edited.";

export function MenuNote({
  id,
  text,
}: {
  id: string;
  text: string;
}): ReactElement {
  return (
    <p id={id} role="none" className={editorClassNames.menuNote}>
      {text}
    </p>
  );
}
