# Comments

## Package structure

A WordprocessingML comment is split between the main document story and a Comments part. The main story carries `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference`; the related Comments part carries the `w:comment` body with the same `w:id`.

The Comments part is related from the main document part with the WordprocessingML `comments` relationship type and declared with the comments content type in `[Content_Types].xml`.

Observed 2026-08-22 against ECMA-376 5th edition, Part 1, §11.3.2 and §§17.13.4.3–17.13.4.6.

## Threads and resolved state

Word stores thread metadata in the `commentsExtended` extension part. A `w15:commentEx` identifies the last paragraph of its `w:comment` through `w15:paraId`; `w15:paraIdParent` identifies the parent comment for a reply, and `w15:done` records whether the root thread is resolved. The matching comment paragraph carries `w14:paraId`.

The editor preserves an untouched extension part byte-identically. Creating a reply or changing resolved state creates or rewrites the part, its main-document relationship, and its content type as needed. A reply added to a resolved thread writes the root as open again. Deleting a root also removes the replies associated with its paragraph id; an unrelated orphan Comments-part entry stays preserved.

Observed 2026-08-22 against Microsoft [MS-DOCX] §2.5.3.1 `CT_CommentEx` and §2.5.1.5 `commentsEx`.

## Anchors

A range comment uses matching start and end markers plus a reference with the same id. A reference without either range marker is a point comment anchored at the reference position. An unmatched start or end marker is also interpreted as a point anchor when a comment reference carries the same id; a marker without that reference is non-conformant. New range comments create the three main-story elements together; removal deletes every supported marker and reference represented for that id.

Imported ranges whose markers occur inside supported paragraphs can cross paragraphs. Markers in block-level or unsupported containers remain inside their preserved enclosing XML and are outside comment editing. New comments are limited to a non-empty selection within one paragraph, which lets the editor place one unambiguous start/end pair without changing surrounding document structures.

## Editing model

The markers are typed, invisible inline nodes. Decorations highlight the text between a matching pair, while comment author, date, initials, and body travel on the reference node and are read from or written to the Comments part.

When comments do not change, their XML and package parts remain byte-identical. If the Comments part is rewritten, unedited entries retain their imported structure and content, but not necessarily their original lexical serialization. Every id already present in the Comments part remains reserved, including an entry not referenced by the main story. Editing a body writes a plain-text `w:comment`; creating the first comment also creates the relationship and content-type override. Deleting through the comment command removes the supported inline range markers, the reference, the matching Comments-part entry, and its replies.
