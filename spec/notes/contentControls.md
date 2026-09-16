# Content controls (`w:sdt`)

## Locking is two obligations, not one

`w:lock` (§17.5.2.23) and the `ST_Lock` values it takes (§17.18.49) settle two independent things: whether the contents may be edited, and whether the control may be deleted in its entirety.
`contentLocked` adds a third clause worth reading twice: the control may be deleted whole, but no sub portion of it may be.

| `w:lock w:val` | Contents editable | Control deletable whole |
| --- | --- | --- |
| `unlocked`, or the element omitted | yes | yes |
| `sdtLocked` | yes | no |
| `contentLocked` | no | yes, and only whole |
| `sdtContentLocked` | no | no |

Both clauses are normative ("shall"), so a document carrying `sdtLocked` is asking for a control whose text anyone may retype and whose wrapper nobody may remove.

Observed 2026-08-20 against ECMA-376 5th edition, Part 1.

## What we implement

`docx/sdt` reads both clauses off the `w:lock` value and carries them apart, as `contentsLocked` and `deletionLocked` on the wrapper, beside `group` for the separate restriction below.
The three travel through the schema as attributes of the inline `sdt` mark and of the `sdtBlock` node, and as `sdtContentsLocked`, `sdtDeletionLocked` and `sdtGroup` on a cell a control wraps.

`schema/locks` judges a step's edited range against each control it meets by how much of the control the range covers.
A range that covers the control from end to end and takes what stands there away is the control being deleted whole, which the deletion clause answers.
Anything less - a partial overlap, an insertion, or a mark laid across the control, which leaves it standing - reaches into the contents, which the contents clause answers.
For a cell the control's extent is the cell node itself, so the range a row or column deletion writes covers it whole.
For a block-level control the extent is the control's own node, so a range that covers that node and takes it away is the control being deleted whole, while any range reaching the blocks inside it is an edit of its contents.
Where one control stands inside another, every control around the spot answers for editing what stands there and any one of them that shuts refuses it, while the outermost control a range covers whole answers for taking it away.

That gives the four values the behavior the table above asks for, including two cases we previously handled incorrectly:

- A `sdtLocked` control keeps its wrapper: a deletion covering it whole is refused, while its contents stay editable.
- A `contentLocked` control may be deleted whole and no less than whole: a partial deletion, a retype or a formatting change is refused.

The lock a `w:lock` states never reaches the way back out: it rides inside the prefix XML the control is preserved as, so a document carrying any of the four goes out byte for byte when nothing was edited.

## A group is shut without saying so, and a lock admits no exception

`w:group` (§17.5.2.17) states that the contents of a group "shall not be editable", whatever its `w:lock` says, and adds in the same paragraph that "this restriction can be superseded by any structured document tag contained within the group".
The `ST_Lock` values (§17.18.49) state no exception of their own, and Word draws a control nested inside a `contentLocked` one as unmodifiable.

So what the lock says and whether the control is a group are two facts a control carries apart, and the decision is that only the group one may be superseded:

- A control whose lock shuts its contents shuts everything inside it, an inner control included: a locked cell holding an open rich text control, a locked block control holding a table, and a locked inline control holding an open inline one all refuse an edit made inside the inner control.
- A group shuts its contents on its own, and a spot standing inside any control the group holds - a block-level control, a wrapped cell, or an inline control inside one of its paragraphs - is open as far as the group is concerned. A lock standing further out still refuses it.
- The deletion clause is the lock's alone: a group says nothing about being removed, so an unlocked group may still be deleted whole.

Where the specification is silent the lock fails closed, since a control that opens itself is not a statement about what stands around it.

A group is therefore the one shut control the editor offers no way to open: lifting a lock takes the `w:lock` away and leaves the `w:group` where it stood, so a control a group alone shuts has no lock to lift, and a selection standing in one reads as shut rather than as a lock in reach.

Observed 2026-09-16.

## A block control is a container, an inline one a mark

A `w:sdt` under `w:body`, inside a `w:tc`, or inside another control's `w:sdtContent` holds blocks rather than inline content (`CT_SdtContentBlock`, §17.5.2.34), so the editor reads it as a node holding those blocks (`docx/importSdtBlock`) rather than as a mark.
Which control stands inside which is then what the tree says, where the inline mark has to carry `depth` as an attribute of its own, since several marks stand on one and the same text.

The control is one block of the story, so a control nobody edited goes back out as the bytes it arrived as and an edit anywhere inside it rewrites the control whole - the same bargain a table makes.

Both shapes read the prefix, what the control states about being edited and deleted, and the copy rule out of `docx/sdt`, so what the four levels - block, inline, cell, row - disagree about is the node, never the vocabulary.

## The id is what keeps a lock from fragmenting

`w:id` (§17.5.2.18) "shall be persisted through multiple sessions (i.e. shall not be changed once specified)".
That settles a question locking raises. Lifting a lock leaves the control standing, so locking a paragraph that once held one meets a control already there and would come back as three: a fresh control on either side of it.
Tidying that into one by dropping the middle control would drop its id, which the specification does not allow.

So the widening goes the other way: where the one control in the stretch says nothing but its own id and its lock, it takes the stretch over and its id travels with it.
Where the control carries an alias, a tag or a `w:dataBinding`, it keeps the stretch it was given, since widening a named control makes the name cover text it never named.
The specification is silent on that second half: it says the id must survive, not how wide a control may grow, so which controls may be widened is ours to decide.

The shape the specification would allow for one lock over the whole paragraph with the inner control intact is nesting, which `w:group` (§17.5.2.17) describes as normal ("This restriction can be superseded by any structured document tag contained within the group").
The editor represents it: a control is a ProseMirror mark that excludes no other, and the depth the file nested it at rides on the mark, so a control read inside another goes back out inside it (`schema/wrappers`).
Locking does not write that shape - what a new lock does is the widening above - but a document that arrives with it keeps it, and a link nested with a control either way round keeps its order too.

Observed 2026-08-20.

## What we write

Locking a stretch from the editor writes `sdtContentLocked`, the strictest of the four.
A part of a contract that has been settled should neither be retyped nor have its wrapper lifted, and the wrapper is what carries the identity a later export has to put back.
The other three values are read and honored but never written.
