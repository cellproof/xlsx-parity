# Changelog

## Milestone 4

47 formula cells in fixtures F11–F25. Fifteen, in fixtures F11–F13, come from a desktop Excel for Microsoft 365 run by Adam Marsh on 2026-11-22. [HyperFormula's list of differences with Microsoft Excel](https://hyperformula.handsontable.com/docs/guide/list-of-differences) is a secondary cross-reference for those formulas; nothing is copied from that page beyond the short formulas. Thirty-two come from IronCalc issues that show Excel's output. Every expected value is copied from its cited source. Strength is strong when the source states the result for that input. None of these 47 is a weaker citation. F04's expected value was later measured as `"0"` by Adam Marsh, desktop Excel (16.0300), on 2026-11-23. The earlier `"-0"` was Univer issue 8270, whose author did not run it in Excel. The MrExcel thread and the Microsoft Q&A on that fixture were inferred support for picking the negative format section from the raw sign before rounding; that guess does not match this formula.

IronCalc 0.9.1, Python and `@ironcalc/nodejs`, agree on every new cell. Python matches Excel on 13 of the 47. The other 34 are recorded misses in [`suite/baseline.json`](suite/baseline.json).

Formulas below are shown as Excel displays them. The stored strings, including `_xlfn.` and `_xlws.`, are in [`fixtures/fixtures.json`](fixtures/fixtures.json).

| Case | Formula | Excel | Source | Strength | IronCalc 0.9.1 |
|---|---|---|---|---|---|
| F11 `Sheet1!A1` | `INT(-8.9)` | -9 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences), [MS INT](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) | strong | -9 |
| F11 `Sheet1!A2` | `ISEVEN(2.5)` | TRUE | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences), [MS ISEVEN](https://support.microsoft.com/en-us/office/iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356) | strong | TRUE |
| F11 `Sheet1!A3` | `ISODD(3.5)` | TRUE | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | TRUE |
| F11 `Sheet1!A4` | `CEILING.MATH(-4.3,2,2)` | -6 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | -6 |
| F11 `Sheet1!A5` | `FLOOR.MATH(-4.7,2,2)` | -4 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | -4 |
| F12 `Sheet1!A1` | `NPV(1,TRUE(),1)` | 0.75 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | 0.5 |
| F12 `Sheet1!A2` | `NPV(1,B1)` with B1 TRUE | 0 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | 0 |
| F12 `Sheet1!A3` | `COUNTA(1,)` | 2 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | 1 |
| F12 `Sheet1!A4` | `T.INV(0,1)` | #DIV/0! | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences). [MS T.INV](https://support.microsoft.com/en-us/office/t-inv-function-2908272b-4e61-4942-9df9-a25fec9b0e2e) says #NUM! for probability <= 0, which is wrong for df=1 | strong | #NUM! |
| F12 `Sheet1!A5` | `SEQUENCE(0)` | #CALC! | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | #CALC! |
| F12 `Sheet1!A6` | `WEIBULL.DIST(0,1,1,FALSE())` | 0 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences); [MS Q&A 5065194](https://learn.microsoft.com/en-us/answers/questions/5065194/weibull-dist-function-error) reports Excel returns 0, and the Microsoft reply does not dispute it | strong | 1 |
| F12 `Sheet1!A7` | `ADDRESS(1,1,4,TRUE(),"")` | !A1 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | !A1 |
| F12 `Sheet1!A8` | `DAYS(-1,0)` | #NUM! | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | #NUM! |
| F12 `Sheet1!A9` | `T.INV(0,2)` | #NUM! | Adam Marsh, desktop Excel 2026-11-22 | strong | #NUM! |
| F13 `Sheet1!A1` | `IF(FALSE(),A1,0)` | 0 | Adam Marsh, desktop Excel 2026-11-22; [HF differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) | strong | 0 |
| F14 `Sheet1!A1` | `LARGE(B1:B4,1.9)` | 5 | [IronCalc #1750](https://github.com/ironcalc/IronCalc/issues/1750) | strong | 8 |
| F14 `Sheet1!A2` | `LARGE(B1:B4,4.5)` | #NUM! | [IronCalc #1750](https://github.com/ironcalc/IronCalc/issues/1750) | strong | 1 |
| F14 `Sheet1!A3` | `LARGE(B1:B4,TRUE)` | 8 | [IronCalc #1750](https://github.com/ironcalc/IronCalc/issues/1750) | strong | #VALUE! |
| F14 `Sheet1!A4` | `SMALL(B1:B4,TRUE)` | 1 | [IronCalc #1750](https://github.com/ironcalc/IronCalc/issues/1750) | strong | #VALUE! |
| F15 `Sheet1!E1` | `VLOOKUP(2,A1:C2,0,FALSE)` | #VALUE! | [IronCalc #1751](https://github.com/ironcalc/IronCalc/issues/1751) | strong | #REF! |
| F15 `Sheet1!E2` | `HLOOKUP(1,A1:C2,0.5,FALSE)` | #VALUE! | [IronCalc #1751](https://github.com/ironcalc/IronCalc/issues/1751) | strong | #REF! |
| F16 `Sheet1!A1` | `TEXT(1/33600,"General")` | 2.97619E-05 | [IronCalc #1752](https://github.com/ironcalc/IronCalc/issues/1752) | strong | 0.000029762 |
| F16 `Sheet1!A2` | `TEXT(12345678.9012345,"General")` | 12345678.9 | [IronCalc #1752](https://github.com/ironcalc/IronCalc/issues/1752) | strong | 12345678.9012345 |
| F16 `Sheet1!A3` | `TEXT(100000000000/3,"General")` | 33333333333 | [IronCalc #1752](https://github.com/ironcalc/IronCalc/issues/1752) | strong | 33333333333.333332062 |
| F17 `Sheet1!A1` | `LEFT("a string","6")` | a stri | [IronCalc #1605](https://github.com/ironcalc/IronCalc/issues/1605) | strong | #VALUE! |
| F17 `Sheet1!A2` | `SUMSQ("2","3",4)` | 29 | [IronCalc #1605](https://github.com/ironcalc/IronCalc/issues/1605) | strong | 16 |
| F17 `Sheet1!A3` | `STDEV("1","2")` | 0.7071067811865476 | [IronCalc #1605](https://github.com/ironcalc/IronCalc/issues/1605), [MS STDEV](https://support.microsoft.com/en-us/office/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0) (text numbers typed as arguments are counted) | strong | #DIV/0! |
| F17 `Sheet1!A4` | `VAR(TRUE,FALSE,TRUE)` | 0.33333333333333337 | [IronCalc #1605](https://github.com/ironcalc/IronCalc/issues/1605), [MS VAR](https://support.microsoft.com/en-us/office/var-function-1f2b7ab2-954d-4e17-ba2c-9e58b15a7da2) (logicals typed as arguments are counted) | strong | #DIV/0! |
| F18 `Sheet1!D1` | `XLOOKUP("b",A1:A3,B1:C3)` | x | [IronCalc #1789](https://github.com/ironcalc/IronCalc/issues/1789) | strong | #VALUE! |
| F18 `Sheet1!G1` | `XLOOKUP(8,E1:E4,F1:F4,,,)` | a | [IronCalc #1718](https://github.com/ironcalc/IronCalc/issues/1718) | strong | #ERROR! |
| F19 `Sheet1!B1` | `INDEX(SORT(A1:A4,,-1),1)` | 8 | [IronCalc #1597](https://github.com/ironcalc/IronCalc/issues/1597) | strong | #VALUE! |
| F19 `Sheet1!B2` | `XMATCH(8,A1:A4,,)` | 3 | [IronCalc #1597](https://github.com/ironcalc/IronCalc/issues/1597) | strong | #VALUE! |
| F20 `Sheet1!A1` | `YEAR(0)` | 1900 | [IronCalc #1569](https://github.com/ironcalc/IronCalc/issues/1569) | strong | #NUM! |
| F20 `Sheet1!A2` | `DAY(0)` | 0 | [IronCalc #1569](https://github.com/ironcalc/IronCalc/issues/1569) | strong | #NUM! |
| F20 `Sheet1!A3` | `WEEKDAY(0)` | 7 | [IronCalc #1569](https://github.com/ironcalc/IronCalc/issues/1569) | strong | #NUM! |
| F20 `Sheet1!A4` | `YEAR(B1)` (B1 blank) | 1900 | [IronCalc #1569](https://github.com/ironcalc/IronCalc/issues/1569) | strong | #NUM! |
| F21 `Sheet1!A1` | `DATE(1,1,1)` | 367 | [IronCalc #1364](https://github.com/ironcalc/IronCalc/issues/1364) | strong | 367 |
| F21 `Sheet1!A2` | `DATE(100,1,1)` | 36526 | [IronCalc #1364](https://github.com/ironcalc/IronCalc/issues/1364) | strong | 36526 |
| F22 `Sheet1!C1` | `AND(A1<>"",A1+B1<1)` | #VALUE! | [IronCalc #1609](https://github.com/ironcalc/IronCalc/issues/1609) | strong | FALSE |
| F23 `Sheet1!E6` | `SUM(IF(TRUE,B1:B4))` | 17 | [IronCalc #1749](https://github.com/ironcalc/IronCalc/issues/1749) | strong | #VALUE! |
| F23 `Sheet1!E7` | `SUM(IF(B1>5,B1:B4,C1:C4))` | 100 | [IronCalc #1749](https://github.com/ironcalc/IronCalc/issues/1749) | strong | #VALUE! |
| F24 `Sheet1!A1` | `SUBTOTAL(6,B1:B2)` | 0 | [IronCalc #1739](https://github.com/ironcalc/IronCalc/issues/1739) | strong | 1 |
| F24 `Sheet1!A8` | `SUBTOTAL(110,H1:H3)` | 2 | [IronCalc #1739](https://github.com/ironcalc/IronCalc/issues/1739) | strong | 9.333333333333334 |
| F24 `Sheet1!A9` | `SUBTOTAL(111,H1:H3)` | 1 | [IronCalc #1739](https://github.com/ironcalc/IronCalc/issues/1739) | strong | 6.222222222222221 |
| F24 `Sheet1!E6` | `SUBTOTAL(9,D3:D6)` | 6 | [IronCalc #1739](https://github.com/ironcalc/IronCalc/issues/1739) | strong | 13 |
| F25 `Sheet1!A1` | `DECIMAL("-10",2)` | #NUM! | [IronCalc #1615](https://github.com/ironcalc/IronCalc/issues/1615) | strong | -2 |
| F25 `Sheet1!A2` | `DECIMAL(-5,10)` | #NUM! | [IronCalc #1615](https://github.com/ironcalc/IronCalc/issues/1615) | strong | -5 |

### Left out

These were not added. The first group has no Excel value stated exactly enough to copy. The second group has a stated value, but IronCalc 0.9.1 rewrites the formula text on save, so the suite's formula-preservation check cannot record them.

No exact Excel value to copy:

- HyperFormula rows whose Excel column is rounded to four places: `AVEDEV(TRUE(),FALSE())` 0.4444, `SKEW(TRUE(),FALSE())` 1.7321, `GAMMA(-2.5)` -0.9453, `VARP(...)` 1.9592, `POISSON.DIST(0,-0.01,FALSE())` 1.0101, `DB(...)` 15845.0985, `TBILLEQ(0,360,0.1)` 0.1097, `TBILLEQ(0,365,0.1)` 0.1098, `HYPGEOM.DIST(4,12,20,40,TRUE())` 0.1504.
- HyperFormula rows that do not state one Excel result: the XLOOKUP empty-cell search ("may be invalid"), `{A1,A2}` ("ERROR: invalid array"), `DEVSQ(A2,A3)` and `CHISQ.TEST` with no cell values, `NORMSDIST` ("Wrong number"), and the broken second dependency-collection example.
- IronCalc #1609's `AND(FALSE,1/0)` and `OR(TRUE,1/0)`. The author says those two rows were not taken from a saved file, and Microsoft's AND page does not state that result. The empty-text row from the same issue is included.

Stated, but the saved formula text would not match:

- `TEXT(0.000000001,"General")` from IronCalc #1752. IronCalc rewrites the literal to `1E-9`.
- `="its "&(2<3)&" blue"` from IronCalc #1605. IronCalc drops the parentheses.
