/**
 * Turning the lock of a content control on and off inside the opening XML it goes back out as.
 *
 * Nothing but the `w:lock` child of the `w:sdtPr` moves. The id Word gave the control, the alias
 * and tag it goes by, and the `w:dataBinding` that ties it to the file's own XML all stay, which
 * is what tells lifting a lock apart from throwing the control away.
 *
 * One child carries both clauses of the lock, so the two move together here: shutting writes the
 * value that shuts both, and lifting takes the child away, which leaves the control with neither.
 * A control Word gave nothing but a guard against deletion (`sdtLocked`) is therefore shut by
 * being given the stricter value, and lifting that lock afterwards opens both clauses.
 */

import { editSdtPrefix } from "./sdt";

/** The value Word writes for a control whose contents may not be edited, and which may not be deleted either */
const LOCK_BOTH_CLAUSES = '<w:lock w:val="sdtContentLocked"/>';

/**
 * The opening of the control with its lock shut or lifted.
 * null for a prefix whose shape cannot be made out, which leaves the caller to decide what to do
 * with a control it cannot rewrite.
 */
export function withContentLock(
  sdtPrefix: string,
  locked: boolean
): string | null {
  return editSdtPrefix(sdtPrefix, [
    ["lock", locked ? LOCK_BOTH_CLAUSES : null],
  ]);
}
