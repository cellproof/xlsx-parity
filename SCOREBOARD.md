# Scoreboard

Last updated: 2026-11-22 (latest `when` in fixtures/fixtures.json).

Expected values are Excel's, copied from a Univer issue, an IronCalc issue, a Microsoft Support page, or a recorded desktop Excel result (F11, F12, and F13 from Adam Marsh, desktop Excel for Microsoft 365, 2026-11-22); HyperFormula's differences page is a reference, not the source of the values. Each miss cites the source stored on that cell, and IronCalc's misses are grouped in [docs/ironcalc-triage.md](docs/ironcalc-triage.md).

Generated from [`suite/baseline.json`](suite/baseline.json) and [`fixtures/fixtures.json`](fixtures/fixtures.json) by `make scoreboard`.

| Engine | Version | Mode | Passing | Percent |
|---|---|---|---|---|
| SheetJS CE | 0.20.3 | read | 74/74 | 100% |
| SheetJS CE | 0.20.3 | write | 74/74 | 100% |
| Univer | 1.6.0 | native | 47/74 | 63.5% |
| LibreOffice Calc | 26.2.5.2 | recalc | 46/74 | 62.2% |
| IronCalc | 0.9.1 | JavaScript | 34/74 | 45.9% |
| IronCalc | 0.9.1 | Python | 34/74 | 45.9% |
| Univer | 1.6.0 | hook | 34/74 | 45.9% |

A cell passes when the recorded value matches Excel. Strings, booleans, and errors are exact. Numbers use a relative tolerance of 1e-12. SheetJS read and write pass when the formula string and the cached value are preserved.

## Misses

### SheetJS CE 0.20.3 read

No misses. The baseline records no SheetJS read failures, so this mode passes every formula cell.

### SheetJS CE 0.20.3 write

No misses. The baseline records no SheetJS write failures, so this mode passes every formula cell.

### Univer 1.6.0 native

27 cells miss.

#### F01-let-lambda-call

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `2` | `#NAME?` | issue #8332 table |
| Sheet1!A2 | `3` | `#NAME?` | issue #8332 table |

#### F03-edate-eomonth-negative-fraction

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `40527` | `"2010/12/15"` | issue #8269 / MS EDATE |

#### F04-text-negative-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"-0"` | `"0"` | issue #8270 (author unverified in Excel; weaker). Format-section support is inferred, not a direct test |

#### F05-search-wildcards-and-round-digits

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B3 | `1.5` | `#NUM!` | issue #8198 text |
| Sheet1!B4 | `0` | `#NUM!` | issue #8198 text |

#### F06-char-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#VALUE!` | `"Ā"` | issue #8211 / MS CHAR |
| Sheet1!A2 | `#VALUE!` | `"𐁁"` | issue #8211 / MS CHAR |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `#SPILL!` | issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `#SPILL!` | issue #8154 text (k299 total) |

#### F11-hf-negative-rounding

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `true` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences / MS ISEVEN |
| Sheet1!A3 | `true` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table (Excel TRUE) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `"$0.75 "` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A2 | `0` | `"$0.00 "` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table (B1 = TRUE) |
| Sheet1!A3 | `2` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |
| Sheet1!A7 | `"!A1"` | `"A1"` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A4 | `0.33333333333333337` | `0.333333333333` | issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#SPILL!` | issue #1789 table (first spill cell) |

#### F21-date-year-offset

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `367` | `"1901/01/01"` | issue #1364 table |
| Sheet1!A2 | `36526` | `"2000/01/01"` | issue #1364 table |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `0` | issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A8 | `2` | `9.33333333333` | issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.22222222222` | issue #1739 table (VAR.P, row 2 hidden) |

### LibreOffice Calc 26.2.5.2 recalc

28 cells miss.

#### F01-let-lambda-call

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `2` | `#VALUE!` | issue #8332 table |
| Sheet1!A2 | `3` | `#VALUE!` | issue #8332 table |

#### F02-iso-ceiling

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `#NAME?` | issue #8268 / MS ISO.CEILING |
| Sheet1!A2 | `-4` | `#NAME?` | issue #8268 / MS ISO.CEILING |
| Sheet1!A3 | `6` | `#NAME?` | issue #8268 / MS ISO.CEILING |

#### F04-text-negative-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"-0"` | `"0"` | issue #8270 (author unverified in Excel; weaker). Format-section support is inferred, not a direct test |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `0` | `0.5` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table (B1 = TRUE) |
| Sheet1!A4 | `#DIV/0!` | `#VALUE!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A5 | `#CALC!` | `#VALUE!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |
| Sheet1!A7 | `"!A1"` | `"A1"` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A8 | `#NUM!` | `#NAME?` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A9 | `#NUM!` | `#VALUE!` | primary: Adam Marsh, desktop Excel 2026-11-22 |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `#NUM!` | `#VALUE!` | issue #1750 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.0000297619047619048"` | issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.3333"` | issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `29` | `#VALUE!` | issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#VALUE!` | issue #1605 table; MS STDEV counts text numbers typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!G1 | `"a"` | `#VALUE!` | issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B2 | `3` | `#VALUE!` | issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `1899` | issue #1569 table |
| Sheet1!A2 | `0` | `30` | issue #1569 table |
| Sheet1!A4 | `1900` | `1899` | issue #1569 table (B1 blank) |

#### F21-date-year-offset

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `367` | `36892` | issue #1364 table |
| Sheet1!A2 | `36526` | `#VALUE!` | issue #1364 table |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `#VALUE!` | issue #1615 table |
| Sheet1!A2 | `#NUM!` | `#VALUE!` | issue #1615 table |

### IronCalc 0.9.1 JavaScript

40 cells miss.

#### F04-text-negative-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"-0"` | `"0"` | issue #8270 (author unverified in Excel; weaker). Format-section support is inferred, not a direct test |

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | issue #1605 table |
| Sheet1!A2 | `29` | `16` | issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | issue #1615 table |

### IronCalc 0.9.1 Python

40 cells miss.

#### F04-text-negative-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"-0"` | `"0"` | issue #8270 (author unverified in Excel; weaker). Format-section support is inferred, not a direct test |

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | issue #1605 table |
| Sheet1!A2 | `29` | `16` | issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | issue #1615 table |

### Univer 1.6.0 hook

40 cells miss.

#### F04-text-negative-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"-0"` | `"0"` | issue #8270 (author unverified in Excel; weaker). Format-section support is inferred, not a direct test |

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: Adam Marsh, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | issue #1605 table |
| Sheet1!A2 | `29` | `16` | issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | issue #1615 table |
