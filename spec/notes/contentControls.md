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

`docx/sdt` reads both clauses off the `w:lock` value and carries them apart, as `contentsLocked` and `deletionLocked` on the wrapper.
They travel through the schema as two attributes of the inline `sdt` mark, and as `sdtContentsLocked` and `sdtDeletionLocked` on a cell a control wraps.

`schema/locks` judges a step's edited range against each control it meets by how much of the control the range covers.
A range that covers the control from end to end and takes what stands there away is the control being deleted whole, which the deletion clause answers.
Anything less - a partial overlap, an insertion, or a mark laid across the control, which leaves it standing - reaches into the contents, which the contents clause answers.
For a cell the control's extent is the cell node itself, so the range a row or column deletion writes covers it whole.

That gives the four values the behavior the table above asks for, including two cases we previously handled incorrectly:

- A `sdtLocked` control keeps its wrapper: a deletion covering it whole is refused, while its contents stay editable.
- A `contentLocked` control may be deleted whole and no less than whole: a partial deletion, a retype or a formatting change is refused.

The lock a `w:lock` states never reaches the way back out: it rides inside the prefix XML the control is preserved as, so a document carrying any of the four goes out byte for byte when nothing was edited.

## The id is what keeps a lock from fragmenting

`w:id` (§17.5.2.18) "shall be persisted through multiple sessions (i.e. shall not be changed once specified)".
That settles a question locking raises. Lifting a lock leaves the control standing, so locking a paragraph that once held one meets a control already there and would come back as three: a fresh control on either side of it.
Tidying that into one by dropping the middle control would drop its id, which the specification does not allow.

So the widening goes the other way: where the one control in the stretch says nothing but its own id and its lock, it takes the stretch over and its id travels with it.
Where the control carries an alias, a tag or a `w:dataBinding`, it keeps the stretch it was given, since widening a named control makes the name cover text it never named.
The specification is silent on that second half: it says the id must survive, not how wide a control may grow, so which controls may be widened is ours to decide.

The shape the specification would allow for one lock over the whole paragraph with the inner control intact is nesting, which `w:group` (§17.5.2.17) describes as normal ("This restriction can be superseded by any structured document tag contained within the group").
The editor cannot represent it: a control is a ProseMirror mark, and two marks of one type cannot nest.

Observed 2026-08-20.

## What we write

Locking a stretch from the editor writes `sdtContentLocked`, the strictest of the four.
A part of a contract that has been settled should neither be retyped nor have its wrapper lifted, and the wrapper is what carries the identity a later export has to put back.
The other three values are read and honored but never written.
