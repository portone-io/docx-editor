---
"@portone/docx-editor": patch
---

A page measurement applies its pushes, break spaces and table continuations in one transaction instead of three. The pages look exactly as they did; what changes is that a single measurement now reaches the editor as a single state change, so anything watching transactions - an `onStateChange` handler, a plugin, a React state hook - sees one rather than three per remeasure.

A table's repeated header now also refreshes as soon as its source row is edited, without waiting for the next measurement.

Changing a continued row's formatting keeps its page gap until remeasurement. Page pushes also update when their measured contribution changes but the total top margin stays the same.
