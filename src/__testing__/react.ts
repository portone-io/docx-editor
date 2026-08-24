/** Drawing a component into a host element, shared across the tests that mount one. */

import { act, type ReactNode } from "react";
import { createRoot } from "react-dom/client";

/**
 * Draws the element into the host and hands back the teardown.
 * Both the drawing and the teardown run inside `act`, so everything React schedules is
 * done by the time either returns
 */
export function renderInto(host: HTMLElement, element: ReactNode): () => void {
  const root = createRoot(host);
  act(() => root.render(element));
  return () => act(() => root.unmount());
}
