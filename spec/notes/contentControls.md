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
The three travel through the schema as attributes of the inline `sdt` mark and of the `sdtBlock` node, and as `sdtContentsLocked`, `sdtDeletionLocked` and `sdtGroup` on a cell or a row a control wraps.

`schema/locks` judges a step's edited range against each control it meets by how much of the control the range covers.
A range that covers the control from end to end and takes what stands there away is the control being deleted whole, which the deletion clause answers.
Anything less - a partial overlap, an insertion, or a mark laid across the control, which leaves it standing - reaches into the contents, which the contents clause answers.
For a cell the control's extent is the cell node itself, so the range a row or column deletion writes covers it whole.
For a row it is the row node, so a row deletion covers it whole while a column deletion takes one cell out of it and reaches its contents instead.
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

A `w:sdt` under `w:tbl` is a `CT_SdtRow` (§17.5.2.30) and rides on the `tableRow` node, exactly as a `CT_SdtCell` rides on the cell.
§17.5.2.35 describes what such a control holds as "a single table row" where `CT_SdtContentRow` admits any number of them, and the stricter reading is the one this editor takes: the row is what carries the wrapper back out, so a control holding two rows has one wrapper and two candidates to hang it on, and the table is kept whole instead.
A control holding another control rather than a row is turned down for the same reason, since the one row inside can carry only one of the two.
That the wrapper is a row's rather than a table's is what the deletion clause is read against: a row deletion takes the control away whole and a column deletion edits what it holds, and a row made beside a wrapped one carries no control at all, since a second control claiming the first one's id is not a shape §17.5.2.18 allows.

## Editing at the edges of a block control

A control names a settled part of a contract, so the join a keystroke builds at its edge may not carry blocks into it or out of it.
That is what an edit meeting the edge is judged by.

- A join at the edge is refused: the keystroke does nothing and the selection stays where it stood.
- A selection running from outside a control into it, or the reverse, is refused whole rather than trimmed to the edge, whatever replaces it: typing, a deletion, a paste.
  Trimming would move the boundary the file drew, and doing nothing is the reading that never carries text into or out of a control behind the user's back.
- A selection covering a control from end to end crosses no edge.
  Taking the control away with everything it held is the deletion clause's question (§17.5.2.23), not this rule's.
- A control left holding a single empty paragraph is removed whole instead, the caret landing where the control stood.
  `w:sdtContent` may hold nothing at all, but Word's own empty control is a paragraph of placeholder text rather than an empty one, so the editor never makes a control with nothing inside it.
  Removing it is judged by the deletion clause like any other whole deletion (§17.5.2.23 `w:lock`), so a `sdtLocked` control refuses it.

The rule is about what a keystroke does on its own, not about what the user asks for.
Text moved out of a control by cutting it and pasting it elsewhere, or by dragging it there, is the user saying where it goes, and is left alone.

Observed 2026-09-16.

## Two properties an edit acts on rather than preserves

Everything else a `w:sdtPr` carries rides back out inside the prefix. These two cannot:

- `w:temporary` (§17.5.2.43) states that the control "shall be removed from the WordprocessingML document when the its contents are modified".
  The first edit inside one therefore takes the wrapper away and leaves the blocks, the runs or the cell it stood around exactly where they were.
- `w:showingPlcHdr` (§17.5.2.39) states that what stands inside is placeholder text rather than contents, and that the state "shall be resumed (showing placeholder text) upon opening this document".
  A file exported with the flag still set over text the user typed would show that text as a placeholder in Word, so the first edit inside drops the flag.
  The placeholder text itself is edited like any other text: nothing selects or replaces it on the user's behalf, and no grey drawing is offered for it.

Both are read in `docx/sdt` beside the locks, so the block, the inline mark and a wrapped cell hear the same thing.
`editor/plugins/controlLifecycle` appends the wrapper change to the edit that caused it, so one undo takes the edit and the wrapper back together, and it runs over a side story as well, so a control inside a footnote settles as one in the body does.

Two silences are settled here.
A control that carries `w:temporary` and a lock against deletion (§17.5.2.23) keeps its wrapper: the lock states that it may not be removed and nothing says which of the two gives way, so the lock wins, the way every silence in this note is settled.
The lock speaks about removing the control and about nothing else, so such a control still drops `w:showingPlcHdr`.
And a change that leaves nothing but a comment is no modification of the contents: a commenter is given the body alone to write against and the file they return is verified against the one they were sent, so a wrapper lifted beside a comment would fail that verification over a change the commenter never made.

Observed 2026-09-16.

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
