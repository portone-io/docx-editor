# Producer lane body text

This is the body every file in `__fixtures__/producers/` carries, so that the same document can be saved out of more than one word processor and the differences between the packages are the software's rather than the prose's.
[README.md](../README.md#producer-lane) says which producer wrote which file and how each one is sanitized.

Type this into the producer and apply the structure named in each section with the producer's own controls, not by editing XML afterwards.
Where a producer offers no equivalent of something below, leave that piece out and record what is missing rather than approximating it.

## Producer lane fixture

Not a demonstration. This file is here so that the importer meets markup a real word processor writes.

*Centred, under the title, in the producer's subtitle size.*

### 1. Why this file is here

Every other fixture in this repository is **built from controlled XML**, so the corpus holds only markup this project chose to write.
This one was saved by a word processor instead, and its text was rewritten afterwards so that nobody reading a failure can mistake it for the readable demo.

A producer stamps a revision identifier on every run and an identifier on every paragraph, writes its own defaults over properties no author set, and reaches for shapes the schema authors did not expect.
None of that is worth **inventing by hand**, and the point of the lane is that the reader survives it.
The paragraph you are reading carries a **suggested edit** left in on purpose, so that the importer meets revision markup it has no model for, and the notes the document opens with say so.

*The suggested edit is a tracked insertion of "Should this sentence say more?" made in the middle of the last sentence, so that it splits a run rather than sitting between two.*

### 2. What the two round trips promise

Opened and exported with no edit at all, every part of this package comes back **byte for byte**.
That promise comes first, because a producer writes properties this reader has no model for and a rebuild would drop them without saying anything.

Editing one paragraph rewrites that paragraph and leaves the rest alone.
The bytes before it and the bytes after it are still the producer's own, and a test compares them against the **original package** rather than against anything this editor produced.

### 3. What this file is known to lose

Three constructs in the body arrive as placeholders rather than as nodes: the suggested edit above, and the pair of markers around the named range further down.
A snapshot beside the tests records all three, so a change in what this document keeps is approved rather than noticed later.

Everything else in the body is modelled: headings, paragraph properties, run formatting, two list definitions, two tables, a link, a comment, a pair of notes and an inline picture.

> **Note.** Parts of what this producer wrote are not valid against the published schemas, and the export hands those bytes back unchanged rather than correcting them. The rejected shapes are listed with the fixture.

*Indented from the left margin and set in a colour of its own, with "Note." in bold.*

### 4. How the sections below were made

Each section was set with the producer's own controls rather than by editing XML, so the markup under it is whatever that software decided to save.

## Markup samples

Every section from here on is named for what it exercises.
The prose under a heading is filler with a purpose: it gives the reader something to rebuild, to demote, or to leave exactly as it stands.

### 5. Alignment

One paragraph per alignment, in this order: left, centred, right, justified.

- This paragraph is aligned left, which is what the producer writes when nothing asked for anything else.
- This paragraph is centred. Alignment belongs to the paragraph, so it travels with the block the exporter rebuilds rather than with a run inside it.
- This paragraph is aligned right. Reading it back has to give the same property, because a lost alignment is a silent change to the page.
- This paragraph is justified. The spacing between the words on every line but the last is stretched to reach both margins, and the property that asks for it survives untouched.

### 6. Indentation

An indent is a paragraph property as well, and the producer writes it in twips.

This paragraph is indented from the left margin.
Whatever value the producer chose is kept as it stands, whether or not this editor would have chosen the same one.

### 7. Line spacing

Line spacing is counted in fractions of a line rather than in points.

One paragraph per spacing, in this order: single, the producer's own default, double.

- This paragraph is single spaced. Its lines sit as close together as the typeface asks for.
- This paragraph is set a little looser than single, which is what the producer applies to a new document on its own.
- This paragraph is double spaced, so the setting is unmistakable beside the two above it.

### 8. Run formatting

Formatting belongs to the run, so one sentence can carry several at once: this stretch is **bold**, this one is *italic*, this one is <u>underlined</u> and ~~this one is struck through~~, each of them a run of its own in the markup.

A size is written in half points, so `this run is smaller` and `this one larger` than the text around them.
Colour and shading are run properties too: `this run is coloured`, and `this one carries a shading fill`.

*In the second paragraph the four marked stretches are, in order, a smaller size, a larger size, a text colour, and a shading fill.*

### 9. Fonts

`This run names a serif face for its Latin characters`, `this one names a monospaced face`, and `this one names one face for Latin text and another for East Asian text`.
Every character in this section is Latin, so that second slot stands unused; it is there to keep a Japanese or Chinese passage out of a face with no glyphs for it.

### 10. Lists

A marker comes from the numbering part, and a list paragraph names a definition and a level inside it.

1. This item sits at the outer level of a numbered definition.
   1. This item sits one level in, where the marker pattern changes.
2. This item is back at the outer level, and it continues the count rather than starting it again.

- This item sits at the outer level of a bulleted definition.
  - This item sits one level in, where the glyph changes with the level.

### 11. A table with merged cells

The rules, the padding and the vertical placement inside a cell are read from the table.
Four of the rows below take part in a merge: one across columns and two down.

Three columns. The first column is a label, the second is aligned left and the third is centred.
The header row is shaded and bold.
The first cell of the third row is merged with the cell below it and holds two paragraphs.
The second cell of the fourth row spans the two columns to its right.
The first cell of the fifth row is merged with the cell below it.

| What the row shows | Left aligned | Centred |
| --- | --- | --- |
| Rules | Every rule in this table is drawn at one weight | and in one colour |
| Widths and padding<br>This cell is merged down | The table width and the cell padding carry decimal points | which the schemas do not accept |
| *merged from above* | This one cell covers the two columns to its right ||
| A merge down | The cell on the left covers this row | and the row below |
| *merged from above* | The first cell of the row above carries the restart marker | and this row carries the continuation |

### 12. A link

The address ending this sentence is a **hyperlink** whose relationship lives beside the body part rather than in it, so a reader has to follow the identifier to learn where it points: [example.com/editor/features](https://example.com/editor/features).

The producer wrote no tooltip on that link and no content control anywhere in the body.
A saved document carrying either is wanted in this lane, and until one arrives neither shape is covered by a file a word processor wrote.

### 13. Comments, bookmarks and notes

This paragraph holds a `commented range` tied to a thread in the comments part.
The producer wrote the reviewer's name into that thread, and the sanitize step replaced it.

A bookmark is an invisible named range. This sentence sits inside one, under a name the producer generated rather than one an author chose.

This sentence carries a footnote and this one carries a second footnote.
Both bodies are in the notes part, and the producer left no endnote part at all.

The header above and the footer below are parts of their own.
The page number in the footer is a field, and the producer wrote the whole of it inside a single run.

*The commented range is the marked stretch, and the comment on it reads: "This thread came out of the producer with a real name on it, which the sanitize step replaced." The bookmark covers the second sentence of the paragraph after it.*

### 14. Cell alignment

The row below is tall enough to show the three vertical placements a cell can ask for.

One row of three cells, each taller than its text, aligned top, centre and bottom in that order.

| Top aligned | Centre aligned | Bottom aligned |
| --- | --- | --- |

### The picture

An inline picture, centred, followed by a centred caption in a smaller italic:

An inline picture stands above this line, drawn at the extent its own drawing records.
Its media part travels inside the package and its relationship is preserved as the producer wrote it.

## The other stories

- Header: `Producer lane fixture`
- Footer: `Page ` followed by a page-number field
- First footnote: ` A footnote sits at the bottom of the page that references it.`
- Second footnote: ` A second footnote, so that the notes part holds more than one.`
