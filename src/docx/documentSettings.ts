import { twipsToPt, wAttr } from "../ooxml/units";
import { childByLocalName } from "../ooxml/xml";

/** The interval OOXML assigns when settings.xml does not declare one. */
export const DEFAULT_TAB_STOP_PT = 36;

/** Reads the document-wide automatic tab interval from settings.xml. */
export function readDefaultTabStop(settings: Document | null): number | null {
  if (!settings) return null;
  const setting = childByLocalName(settings.documentElement, "defaultTabStop");
  const value = setting ? twipsToPt(wAttr(setting, "val")) : null;
  return value !== null && value > 0 ? value : null;
}
