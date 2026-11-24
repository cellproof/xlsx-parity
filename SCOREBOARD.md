# Scoreboard

Last updated: 2026-11-23 (latest `when` in fixtures/fixtures.json).

Expected values come from recorded desktop Excel runs, plus the remaining cited sources. HyperFormula's differences page is a reference, not the source of the values. Each miss cites the source stored on that cell, and IronCalc's misses are grouped in [docs/ironcalc-triage.md](docs/ironcalc-triage.md).

Generated from [`suite/baseline.json`](suite/baseline.json) and [`fixtures/fixtures.json`](fixtures/fixtures.json) by `make scoreboard`.

| Engine | Version | Mode | Passing | Percent |
|---|---|---|---|---|
| SheetJS CE | 0.20.3 | read | 74/74 | 100% |
| SheetJS CE | 0.20.3 | write | 74/74 | 100% |
| Univer | 1.6.0 | native | 49/74 | 66.2% |
| LibreOffice Calc | 26.2.5.2 | recalc | 47/74 | 63.5% |
| IronCalc | 0.9.1 | JavaScript | 35/74 | 47.3% |
| IronCalc | 0.9.1 | Python | 35/74 | 47.3% |
| Univer | 1.6.0 | hook | 35/74 | 47.3% |

A cell passes when the recorded value matches Excel. Strings, booleans, and errors are exact. Numbers use a relative tolerance of 1e-12. SheetJS read and write pass when the formula string and the cached value are preserved.

The Univer hook is IronCalc running inside Univer. SheetJS CE does not calculate, so its row is a read/write (I/O) check.

## Provenance

Cell counts by source type. Classification uses the text before `; previously:`: a `primary: adamlocalis` prefix is adamlocalis's desktop Excel run; a source starting with `MS ` is Microsoft docs; a source that names an Excel run or an Excel-saved workbook is a third-party report of an Excel run; anything else is an issue-stated value.

| Source type | Cells |
|---|---|
| adamlocalis's desktop Excel run | 74 |
| third-party report of an Excel run | 0 |
| issue-stated value | 0 |
| Microsoft docs | 0 |

## Misses

### SheetJS CE 0.20.3 read

No misses. The baseline records no SheetJS read failures, so this mode passes every formula cell.

### SheetJS CE 0.20.3 write

No misses. The baseline records no SheetJS write failures, so this mode passes every formula cell.

### Univer 1.6.0 native

25 cells miss.

#### F01-let-lambda-call

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `2` | `#NAME?` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8332 table |
| Sheet1!A2 | `3` | `#NAME?` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8332 table |

#### F03-edate-eomonth-negative-fraction

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `40527` | `"2010/12/15"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8269 / MS EDATE |

#### F05-search-wildcards-and-round-digits

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B3 | `1.5` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8198 text |
| Sheet1!B4 | `0` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8198 text |

#### F06-char-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#VALUE!` | `"Ā"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8211 / MS CHAR |
| Sheet1!A2 | `#VALUE!` | `"𐁁"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8211 / MS CHAR |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `#SPILL!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `#SPILL!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k299 total) |

#### F11-hf-negative-rounding

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `true` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences / MS ISEVEN |
| Sheet1!A3 | `true` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table (Excel TRUE) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `"$0.75 "` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A2 | `0` | `"$0.00 "` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table (B1 = TRUE) |
| Sheet1!A3 | `2` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |
| Sheet1!A7 | `"!A1"` | `"A1"` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A4 | `0.33333333333333337` | `0.333333333333` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS VAR counts logicals typed as arguments |

#### F21-date-year-offset

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `367` | `"1901/01/01"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1364 table |
| Sheet1!A2 | `36526` | `"2000/01/01"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1364 table |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A8 | `2` | `9.33333333333` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.22222222222` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.P, row 2 hidden) |

### LibreOffice Calc 26.2.5.2 recalc

27 cells miss.

#### F01-let-lambda-call

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `2` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8332 table |
| Sheet1!A2 | `3` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8332 table |

#### F02-iso-ceiling

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `#NAME?` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8268 / MS ISO.CEILING |
| Sheet1!A2 | `-4` | `#NAME?` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8268 / MS ISO.CEILING |
| Sheet1!A3 | `6` | `#NAME?` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8268 / MS ISO.CEILING |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `0` | `0.5` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table (B1 = TRUE) |
| Sheet1!A4 | `#DIV/0!` | `#VALUE!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A5 | `#CALC!` | `#VALUE!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |
| Sheet1!A7 | `"!A1"` | `"A1"` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A8 | `#NUM!` | `#NAME?` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A9 | `#NUM!` | `#VALUE!` | primary: adamlocalis, desktop Excel 2026-11-22 |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `#NUM!` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.0000297619047619048"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.3333"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A2 | `29` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS STDEV counts text numbers typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!G1 | `"a"` | `#VALUE!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B2 | `3` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `1899` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A2 | `0` | `30` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A4 | `1900` | `1899` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table (B1 blank) |

#### F21-date-year-offset

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `367` | `36892` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1364 table |
| Sheet1!A2 | `36526` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1364 table |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |
| Sheet1!A2 | `#NUM!` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |

### IronCalc 0.9.1 JavaScript

39 cells miss.

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A2 | `29` | `16` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |

### IronCalc 0.9.1 Python

39 cells miss.

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A2 | `29` | `16` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |

### Univer 1.6.0 hook

39 cells miss.

#### F07-value-accounting-negatives

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `-5` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A2 | `-1000` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |
| Sheet1!A3 | `-4` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8231 text |

#### F10-sumifs-legacy-whole-column

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Summary!B2 | `10809` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k0 total) |
| Summary!B301 | `10800` | `0` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #8154 text (k299 total) |

#### F12-hf-npv-sequence-errors

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0.75` | `0.5` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A3 | `2` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A4 | `#DIV/0!` | `#NUM!` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table |
| Sheet1!A6 | `0` | `1` | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 (Excel returns 0, reply does not dispute it) |

#### F14-large-k

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `5` | `8` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A2 | `#NUM!` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A3 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |
| Sheet1!A4 | `1` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1750 table |

#### F15-vlookup-index-below-one

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E1 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |
| Sheet1!E2 | `#VALUE!` | `#REF!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1751 table |

#### F16-text-general

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"2.97619E-05"` | `"0.000029762"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A2 | `"12345678.9"` | `"12345678.9012345"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |
| Sheet1!A3 | `"33333333333"` | `"33333333333.333332062"` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1752 table |

#### F17-explicit-argument-coercion

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `"a stri"` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A2 | `29` | `16` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table |
| Sheet1!A3 | `0.7071067811865476` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS STDEV counts text numbers typed as arguments |
| Sheet1!A4 | `0.33333333333333337` | `#DIV/0!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1605 table; MS VAR counts logicals typed as arguments |

#### F18-xlookup-return-and-empty

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!D1 | `"x"` | `#VALUE!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1789 table (first spill cell) |
| Sheet1!G1 | `"a"` | `#ERROR!` | primary: adamlocalis, desktop Excel, 2026-11-23, rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt; previously: issue #1718 table |

#### F19-sort-empty-argument

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!B1 | `8` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |
| Sheet1!B2 | `3` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1597 table |

#### F20-date-serial-zero

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A2 | `0` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A3 | `7` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table |
| Sheet1!A4 | `1900` | `#NUM!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1569 table (B1 blank) |

#### F22-and-empty-text

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!C1 | `#VALUE!` | `false` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1609 table |

#### F23-sum-if-range

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!E6 | `17` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |
| Sheet1!E7 | `100` | `#VALUE!` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1749 table |

#### F24-subtotal-edges

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `0` | `1` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (product of no numbers) |
| Sheet1!A8 | `2` | `9.333333333333334` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.S, row 2 hidden) |
| Sheet1!A9 | `1` | `6.222222222222221` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (VAR.P, row 2 hidden) |
| Sheet1!E6 | `6` | `13` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1739 table (nested SUBTOTAL) |

#### F25-decimal-leading-sign

| Cell | Expected | Actual | Source |
|---|---|---|---|
| Sheet1!A1 | `#NUM!` | `-2` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |
| Sheet1!A2 | `#NUM!` | `-5` | primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously: issue #1615 table |
