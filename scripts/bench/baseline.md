# Scaling baseline

Measured on Apple M4 with 16 GiB RAM, macOS 26.5.1, and Node.js 22.18.0. Times are milliseconds; ratios compare each size with the preceding row. Each case discards a warm-up pass before its first timed size.

| Case | Paragraphs | Import | State | Edit | Export |
| --- | ---: | ---: | ---: | ---: | ---: |
| plain | 500 | 93.4 | 48.5 | 1.5 | 144.5 |
| plain | 1,000 | 175.1 (1.87x) | 96.4 (1.99x) | 5.0 (3.27x) | 549.5 (3.80x) |
| plain | 2,000 | 336.1 (1.92x) | 203.5 (2.11x) | 18.9 (3.75x) | 2,303.6 (4.19x) |
| plain | 4,000 | 765.7 (2.28x) | 525.3 (2.58x) | 76.5 (4.05x) | 13,887.5 (6.03x) |
| rich | 500 | 124.0 | 50.2 | 2.3 | 303.3 |
| rich | 1,000 | 232.7 (1.88x) | 128.4 (2.56x) | 7.6 (3.32x) | 1,297.7 (4.28x) |
| rich | 2,000 | 450.5 (1.94x) | 222.6 (1.73x) | 24.9 (3.30x) | 5,274.7 (4.06x) |
| rich | 4,000 | 973.7 (2.16x) | 552.2 (2.48x) | 99.2 (3.98x) | 62,668.2 (11.88x) |

The export column is load-sensitive, swinging several times over between a busy and an idle machine, so only ratios measured on an idle machine are comparable with these.

## Notes

The notes case opens 1,000 paragraphs referring to 280 footnotes and 20 endnotes, each note holding a bold run and every fourth a second paragraph, beside the same paragraphs referring to nothing.
`Markup` draws every note's story as the notes around the page draw it.
`Demands` asks the footnote source about every block once, and `Again` asks a second time, as the next layout pass does; the source remembers what each block node holds, and in jsdom nothing has a layout to read.
`Layout` is the median of 21 page layout passes over synthetic 20px blocks carrying those demands, with the footnote band for the document with notes and no band for the one without, which is what the editor hands over.
`Body key` is the median of 21 keystrokes in the body, and `Note key` the median of 21 keystrokes inside a footnote, which go in as the story change the view over a note writes.
A keystroke in a note costs what one in the body costs: both run the same plugins over the same document, and a note's is one document-attribute step rather than a step inside the text.
Times are milliseconds, measured on the same machine as above.

| Case | Import | State | Markup | Demands | Again | Layout | Body key | Note key |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| no notes | 34.4 | 10.1 | 0.0 | 0.23 | 0.02 | 0.25 | 11.1 | - |
| 280 footnotes, 20 endnotes | 90.9 | 28.8 | 15.0 | 1.10 | 0.50 | 0.22 | 21.2 | 23.3 |
