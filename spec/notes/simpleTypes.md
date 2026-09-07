# Simple types

A measurement and a boolean each have more than one spelling in WordprocessingML, and an element that records a value twice over says which of the two spellings wins.
All three are properties of the format rather than of any one property element, so a reader that decides them at the call site decides them differently from the next reader.

## A measurement is a union, not a count

`ST_TwipsMeasure` (§22.9.2.14) unites `ST_UnsignedDecimalNumber` with `ST_PositiveUniversalMeasure` (§22.9.2.12), whose lexical space is `[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)`.
So `w:pgSz w:w="8.5in"` and `w:pgSz w:w="12240"` (§17.6.13) name the same paper, and a reader taking the leading digits of the first reads a page 8.5 twips wide.
The same union stands behind every measurement the format writes: `ST_SignedTwipsMeasure` (§17.18.81) for `w:ind`, `w:tab/@pos` and `w:spacing/@line`, `ST_HpsMeasure` (§17.18.42) for `w:sz`, `ST_SignedHpsMeasure` (§17.18.80), and `ST_MeasurementOrPercent` (§17.18.107) for a table or cell width, which admits a percentage besides.
A universal measure is absolute, so it says the same length whatever unit the attribute otherwise counts in: an inch is 1440 twips, 144 half-points and 576 eighths of a point alike, and a pica and a pi are both twelve points.

`ST_EighthPointMeasure` (§17.18.23) and `ST_UnsignedDecimalNumber` (§22.9.2.16) are restrictions rather than unions, so those two count and never carry a unit.
Signed integer counts allow an optional `+` sign ([XML Schema Part 2 §3.3.13](https://www.w3.org/TR/xmlschema-2/#integer)). XML Schema 1.0 `unsignedLong` excludes it (§3.3.21.1); unsigned readers retain the previous parser's tolerance for it as an import compatibility policy. Universal measures and percentages have separate lexical patterns; this does not admit `+1in`.

### Producers write a fraction where the schema counts whole

`ST_DecimalNumberOrPercent` and the unsigned counts are integer types, but the documents a word processor saves do not always keep to that.
Google Docs writes `w:tblW w:w="9026.0"` and cell margins in the same shape, and Word opens those files.
A measurement reader retains the fraction; a count reader such as `ST_DecimalNumber` truncates it for compatibility with the previous numbering reader. This is tolerant import behavior, not schema validation.
Integer formatters reject fractions. Writers that convert points or rebuild table widths and grid columns explicitly round to whole twips (or fiftieths for percentages). Thus editing a table can normalize its width spelling; untouched XML remains preserved.

### An explicit percentage takes precedence over the width type

Word interprets `w:w="50%"` as 50 percent even when `w:type` names `dxa`, `auto`, or `nil`, according to [MS-OI29500 §2.1.185(b)](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d9b4a6f6-8b07-43c5-87f2-34c1238911a4).
Table and cell readers follow that rule and store 2500 fiftieths. Bare counts still use `w:type`; a universal measure paired with `pct` remains unreadable because this project has no supported interpretation for that combination.

## A boolean has six spellings and one of them is silence

`ST_OnOff` (§22.9.2.7) unites `xsd:boolean` with the pair `on` and `off`, so `1`, `true` and `on` all state on, and `0`, `false` and `off` all state off.
The union is case-sensitive: `On` and `TRUE` are neither.
A boolean property element that carries no `w:val` at all states on (§17.17.4), which is why `<w:b/>` is bold, and a property element that is not there states nothing.

For an invalid or empty `w:val`, this project uses the present element's default of on as a recovery policy; this is not a guarantee about Word's behavior. An invalid, empty, or absent `w:default` does not designate a default style, so it remains false.
`w:default` on a style is `ST_OnOff` like any other, so a style marked `w:default="on"` is the default style for its kind exactly as `w:default="1"` is.

## Writing one attribute of a pair means dropping the other

Several elements record the same setting twice over, and the format says which of the pair a consumer reads.

| Element | Attribute | Overridden by |
| --- | --- | --- |
| `w:ind` (§17.3.1.12) | `w:left`, `w:start`, `w:right`, `w:end`, `w:hanging`, `w:firstLine` | the `*Chars` attribute of the same name, which counts the indent in characters |
| `w:rFonts` (§17.3.2.26) | `w:ascii`, `w:hAnsi`, `w:eastAsia`, `w:cs` | `w:asciiTheme`, `w:hAnsiTheme`, `w:eastAsiaTheme`, `w:cstheme` |
| `w:color` (§17.3.2.6) | `w:val` | `w:themeColor`, with `w:themeTint` and `w:themeShade` |
| `w:shd` (§17.3.5) | `w:color`, `w:fill` | `w:themeColor` and `w:themeFill`, each with its tint and shade |
| `w:spacing` (§17.3.1.33) | `w:before`, `w:after` | `w:beforeLines`, `w:afterLines` |

Writing one of a pair and leaving the other standing writes a value the consumer will not show, so the overriding attribute is dropped as the overridden one is written.
The reverse direction is a separate question: resolving a theme color or a measurement in character units needs the theme part and the document grid, and until a reader has those it reads the attribute it can and leaves the other in the preserved XML.
