/** Validates document-provided addresses before the link card hands them to the browser. */

const OPENABLE_SCHEMES = ["http:", "https:", "mailto:"];

/**
 * Whether this address is one to hand a browser.
 *
 * A scheme-less address is not, and a leading `www.` does not make it one: nothing in the URI rules
 * says a bare host means https, and a relationship target with no scheme is read as relative to the
 * package, which is to say a file beside the document rather than a site. Guessing here would be
 * guessing what the document meant, which is `setLink`'s policy not to do.
 */
export function openable(address: string): boolean {
  try {
    return OPENABLE_SCHEMES.includes(new URL(address).protocol);
  } catch {
    return false;
  }
}

export function openAddress(address: string): void {
  window.open(address, "_blank", "noopener,noreferrer");
}
