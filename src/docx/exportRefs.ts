/**
 * What the serializers need from the rest of the package while a body is being written: which
 * relationship an inserted image ended up on (`docx/media`), which one each link's address lives
 * on (`docx/hyperlink`), and where to record an approximation a writer had to make
 * (`docx/fidelity`).
 *
 * They travel as one value so that a serializer taking a body apart passes on everything the
 * pieces of it may need, whatever a later kind of content turns out to require.
 */

import { type FidelityCollector, NO_FIDELITY_COLLECTOR } from "./fidelity";
import { type LinkRefs, NO_LINK_REFS } from "./hyperlink";
import { type ImageRefs, NO_IMAGE_REFS } from "./media";

export interface ExportRefs {
  images: ImageRefs;
  links: LinkRefs;
  notes: FidelityCollector;
}

/** What a serializer called on its own, outside an export, has to work with */
export const NO_EXPORT_REFS: ExportRefs = {
  images: NO_IMAGE_REFS,
  links: NO_LINK_REFS,
  notes: NO_FIDELITY_COLLECTOR,
};
