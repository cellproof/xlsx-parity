# IronCalc 0.9.1: 39 misses against Excel

IronCalc 0.9.1 passes 35 of 74 formula cells in [`suite/baseline.json`](../suite/baseline.json). The other 39 are the same on the Python bindings and on `@ironcalc/nodejs`. This note groups those 39 by root cause. It does not change the baseline.

Source read:

- Tag `v0.9.1` is `eb9ca75bd99160c17902359d2d87225c4571c468` (2026-10-21, "UPDATE: Prepare for version v0.9.1").
- `main` at the time of this note is `f93662006b8fb1c401dc870bbda26c5d60ac6037` (2026-11-20, "FIX: a sheet whose r:id points at no relationship is an error, not a panic"). Cargo.toml on that commit still says version 0.9.1.

How `main` was checked: the repo was cloned, built with rustc 1.100.0, and each fixture was loaded with `ironcalc::import::load_from_xlsx`, recalculated with `Model::evaluate`, and read with `get_cell_value_by_index`. Every fixture imported as `DateSystem::Excel1900`, because xlsxwriter writes `Application` as `Microsoft Excel` and `date_system_for` in `xlsx/src/import/mod.rs` treats that as Excel 1900. The "main" column below is that run. Nothing was pushed to IronCalc.

Expected values and citations are copied from [`fixtures/fixtures.json`](../fixtures/fixtures.json). The citation column below is the previous citation. Cells remeasured on 2026-11-23 are prefixed there with `primary: adamlocalis, desktop Excel (16.0300), 2026-11-23; previously:` (F18 D1 and G1 cite the rebuilt fixture file opened directly: D1 = x, E1 = y (spilled), G1 = a, no repair prompt). F11–F13 already name the 2026-11-22 run. Formulas are the Excel display form. Stored `_xlfn.` / `_xlws.` prefixes are in that file.

## Five fixes to try first

These five are still wrong on `main`, have no open pull request that makes the change, and are a few lines in one or two functions. Each one was applied in a local scratch clone of `f936620` (not pushed). After the change, the fixture cell matched Excel. IronCalc's own tests on that tree:

- `cargo test -p ironcalc_base --lib`: 2961 passed, 0 failed, 8 ignored. The ignored tests were already ignored; the scratch diff adds no `#[ignore]`.
- `cargo test -p ironcalc --tests`: every integration binary passed, including `generated_tests` (263 passed, 0 failed). That run is the workbook suite under `xlsx/tests/`, load and save.

| Order | Cells | Change | Result |
|---|---|---|---|
| 1 | F12 `Sheet1!A3` | `fn_counta` counts `EmptyArg` | Proven. `COUNTA(1,)` is 2. |
| 2 | F12 `Sheet1!A4` | `fn_t_inv` returns `#DIV/0!` for probability 0 and 1 degree of freedom | Proven. `T.INV(0,1)` is `#DIV/0!`. `T.INV(0,10)` in the existing smoke test stays `#NUM!`. |
| 3 | F14 `A1`–`A4` | `fn_large` ceilings `k`; `fn_large` and `fn_small` accept a boolean `k` | Proven. 5, `#NUM!`, 8, 1. |
| 4 | F23 `E6`, `E7` | `IF`'s result arguments are vectors, so xlsx import does not wrap them in `@` | Proven. 17 and 100. The loaded formulas no longer contain `@`. |
| 5 | F10 `Summary!B2`, `B301` | `SUMIFS` criteria arguments are scalars, so import intersects `$A:$A` | Proven. 10809 and 10800. The loaded formula is `SUMIFS(Data!B:B,Data!A:A,@$A:$A)`. |

One other one-function candidate was left untried. Its size below is a guess: `fn_weibull_dist` returning 0 for a pdf at `x = 0` (F12 `A6`). F04 `A1` is not a candidate. The old expected `"-0"` was wrong: adamlocalis's desktop Excel (16.0300) run on 2026-11-23 returned `"0"`, which is what IronCalc already returns. Open pull requests [#1754](https://github.com/ironcalc/IronCalc/pull/1754) and [#1697](https://github.com/ironcalc/IronCalc/pull/1697) edit the same `is_negative` expression in the digit formatter and still leave a rounded zero unsigned, so on this cell they match Excel.

### Diff that was tried

`fn_counta` in `base/src/functions/statistical/count_and_average.rs`:

```rust
// An omitted argument (COUNTA(1,)) counts. An empty cell does not.
CalcResult::EmptyArg => {
    result += 1.0;
}
CalcResult::EmptyCell => {}
```

`fn_t_inv` in `base/src/functions/statistical/t_dist.rs`, after `df` is truncated and before the existing `#NUM!` check:

```rust
// T.INV(0, 1) is the Cauchy inverse at 0. Excel returns #DIV/0! for that
// pair and #NUM! for every other probability outside (0, 1), including
// T.INV(0, 2). df is already truncated, so 1.9 is treated as 1.
if p == 0.0 && df == 1.0 {
    return CalcResult::Error {
        error: Error::DIV,
        origin: cell,
        message: "T.INV is undefined for probability 0 and 1 degree of freedom".to_string(),
    };
}
```

`fn_large` and `fn_small` in the same file as `fn_counta`. `LARGE` uses `ceil`; `SMALL` keeps `trunc`. Both call `get_number` instead of `get_number_no_bools`, so `TRUE` is 1.

```rust
// fn_large
let k = match self.get_number(&args[1], cell) {
    Ok(f) => f.ceil(),
    Err(s) => return s,
};

// fn_small
let k = match self.get_number(&args[1], cell) {
    Ok(f) => f.trunc(),
    Err(s) => return s,
};
```

`get_function_args_signature` in `base/src/expressions/parser/static_analysis.rs`. Import calls `add_implicit_intersection` from `xlsx/src/import/worksheets.rs`. A `Scalar` argument that is a range is wrapped in `@`. A `Vector` argument is not.

```rust
fn args_signature_if(arg_count: usize) -> Vec<Signature> {
    if !(2..=3).contains(&arg_count) {
        return vec![Signature::Error; arg_count];
    }
    let mut result = vec![Signature::Vector; arg_count];
    result[0] = Signature::Scalar;
    result
}

fn args_signature_sumifs(arg_count: usize) -> Vec<Signature> {
    let mut result = vec![Signature::Vector; arg_count];
    if arg_count >= 3 && !arg_count.is_multiple_of(2) {
        let mut i = 2;
        while i < arg_count {
            result[i] = Signature::Scalar;
            i += 2;
        }
    }
    result
}
```

`Function::If` calls `args_signature_if`. `Function::Sumifs` calls `args_signature_sumifs`. The parser test `simple_test` expected `IF(@A1:A10,@B1:B10,@C1:C10)` and now expects `IF(@A1:A10,B1:B10,C1:C10)`, plus `SUMIFS(B:B,A:A,$A:$A)` importing as `SUMIFS(B:B,A:A,@$A:$A)`. Excel export of those formulas still drops the operator, which is what that test also checks.

## Cells

"main" is the value measured on `f936620` before the scratch diff. "Size" is a guess except for the five groups above, which were edited.

| Group | Fixture | Cell | Formula | Excel | Citation in fixtures.json | IronCalc 0.9.1 | main |
|---|---|---|---|---|---|---|---|
| 1 | F14-large-k | Sheet1!A1 | `LARGE(B1:B4,1.9)` | 5 | issue #1750 table | 8 | 8 |
| 1 | F14-large-k | Sheet1!A2 | `LARGE(B1:B4,4.5)` | #NUM! | issue #1750 table | 1 | 1 |
| 1 | F14-large-k | Sheet1!A3 | `LARGE(B1:B4,TRUE)` | 8 | issue #1750 table | #VALUE! | #VALUE! |
| 1 | F14-large-k | Sheet1!A4 | `SMALL(B1:B4,TRUE)` | 1 | issue #1750 table | #VALUE! | #VALUE! |
| 2 | F12-hf-npv-sequence-errors | Sheet1!A3 | `COUNTA(1,)` | 2 | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table | 1 | 1 |
| 3 | F12-hf-npv-sequence-errors | Sheet1!A4 | `T.INV(0,1)` | #DIV/0! | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table | #NUM! | #NUM! |
| 4 | F23-sum-if-range | Sheet1!E6 | `SUM(IF(TRUE,B1:B4))` | 17 | issue #1749 table | #VALUE! | #VALUE! |
| 4 | F23-sum-if-range | Sheet1!E7 | `SUM(IF(B1>5,B1:B4,C1:C4))` | 100 | issue #1749 table | #VALUE! | #VALUE! |
| 5 | F10-sumifs-legacy-whole-column | Summary!B2 | `SUMIFS(Data!B:B,Data!A:A,$A:$A)` | 10809 | issue #8154 text (k0 total) | 0 | 0 |
| 5 | F10-sumifs-legacy-whole-column | Summary!B301 | `SUMIFS(Data!B:B,Data!A:A,$A:$A)` | 10800 | issue #8154 text (k299 total) | 0 | 0 |
| 6 | F15-vlookup-index-below-one | Sheet1!E1 | `VLOOKUP(2,A1:C2,0,FALSE)` | #VALUE! | issue #1751 table | #REF! | #VALUE! |
| 6 | F15-vlookup-index-below-one | Sheet1!E2 | `HLOOKUP(1,A1:C2,0.5,FALSE)` | #VALUE! | issue #1751 table | #REF! | #VALUE! |
| 7 | F19-sort-empty-argument | Sheet1!B1 | `INDEX(SORT(A1:A4,,-1),1)` | 8 | issue #1597 table | #VALUE! | 8 |
| 7 | F19-sort-empty-argument | Sheet1!B2 | `XMATCH(8,A1:A4,,)` | 3 | issue #1597 table | #VALUE! | 3 |
| 8 | F20-date-serial-zero | Sheet1!A1 | `YEAR(0)` | 1900 | issue #1569 table | #NUM! | 1900 |
| 8 | F20-date-serial-zero | Sheet1!A2 | `DAY(0)` | 0 | issue #1569 table | #NUM! | 0 |
| 8 | F20-date-serial-zero | Sheet1!A3 | `WEEKDAY(0)` | 7 | issue #1569 table | #NUM! | 7 |
| 8 | F20-date-serial-zero | Sheet1!A4 | `YEAR(B1)` (B1 blank) | 1900 | issue #1569 table (B1 blank) | #NUM! | 1900 |
| 9 | F22-and-empty-text | Sheet1!C1 | `AND(A1<>"",A1+B1<1)` | #VALUE! | issue #1609 table | FALSE | #VALUE! |
| 10 | F25-decimal-leading-sign | Sheet1!A1 | `DECIMAL("-10",2)` | #NUM! | issue #1615 table | -2 | #NUM! |
| 10 | F25-decimal-leading-sign | Sheet1!A2 | `DECIMAL(-5,10)` | #NUM! | issue #1615 table | -5 | #NUM! |
| 11 | F24-subtotal-edges | Sheet1!A1 | `SUBTOTAL(6,B1:B2)` | 0 | issue #1739 table (product of no numbers) | 1 | 0 |
| 11 | F24-subtotal-edges | Sheet1!A8 | `SUBTOTAL(110,H1:H3)` | 2 | issue #1739 table (VAR.S, row 2 hidden) | 9.333333333333334 | 2 |
| 11 | F24-subtotal-edges | Sheet1!A9 | `SUBTOTAL(111,H1:H3)` | 1 | issue #1739 table (VAR.P, row 2 hidden) | 6.222222222222221 | 1 |
| 12 | F24-subtotal-edges | Sheet1!E6 | `SUBTOTAL(9,D3:D6)` | 6 | issue #1739 table (nested SUBTOTAL) | 13 | 13 |
| 13 | F16-text-general | Sheet1!A1 | `TEXT(1/33600,"General")` | 2.97619E-05 | issue #1752 table | 0.000029762 | 0.000029762 |
| 13 | F16-text-general | Sheet1!A2 | `TEXT(12345678.9012345,"General")` | 12345678.9 | issue #1752 table | 12345678.9012345 | 12345678.9012345 |
| 13 | F16-text-general | Sheet1!A3 | `TEXT(100000000000/3,"General")` | 33333333333 | issue #1752 table | 33333333333.333332062 | 33333333333.333332062 |
| 14 | F17-explicit-argument-coercion | Sheet1!A1 | `LEFT("a string","6")` | a stri | issue #1605 table | #VALUE! | #VALUE! |
| 14 | F17-explicit-argument-coercion | Sheet1!A2 | `SUMSQ("2","3",4)` | 29 | issue #1605 table | 16 | 16 |
| 14 | F17-explicit-argument-coercion | Sheet1!A3 | `STDEV("1","2")` | 0.7071067811865476 | issue #1605 table; MS STDEV counts text numbers typed as arguments | #DIV/0! | #DIV/0! |
| 14 | F17-explicit-argument-coercion | Sheet1!A4 | `VAR(TRUE,FALSE,TRUE)` | 0.33333333333333337 | issue #1605 table; MS VAR counts logicals typed as arguments | #DIV/0! | #DIV/0! |
| 15 | F18-xlookup-return-and-empty | Sheet1!D1 | `XLOOKUP("b",A1:A3,B1:C3)` | x | issue #1789 table (first spill cell) | #VALUE! | #VALUE! |
| 16 | F18-xlookup-return-and-empty | Sheet1!G1 | `XLOOKUP(8,H1:H4,I1:I4,,,)` | a | issue #1718 table | #ERROR! | #ERROR! |
| 17 | F12-hf-npv-sequence-errors | Sheet1!A1 | `NPV(1,TRUE(),1)` | 0.75 | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table | 0.5 | 0.5 |
| 18 | F12-hf-npv-sequence-errors | Sheet1!A6 | `WEIBULL.DIST(0,1,1,FALSE)` | 0 | primary: adamlocalis, desktop Excel 2026-11-22; HF differences table; MS Q&A 5065194 | 1 | 1 |
| 20 | F07-value-accounting-negatives | Sheet1!A1 | `VALUE("(5)")` | -5 | issue #8231 text | #VALUE! | #VALUE! |
| 20 | F07-value-accounting-negatives | Sheet1!A2 | `VALUE("($1,000)")` | -1000 | issue #8231 text | #VALUE! | #VALUE! |
| 20 | F07-value-accounting-negatives | Sheet1!A3 | `"(5)"+1` | -4 | issue #8231 text | #VALUE! | #VALUE! |

F12 `A2` `NPV(1,B1)` with B1 TRUE is 0 in Excel and in IronCalc 0.9.1, so it is not in the failure list. F12 `A5`, `A7`, `A8`, and `A9` also match and are omitted here. F04 `TEXT(-0.4,"0")` matches too: Excel and IronCalc both return `"0"`. The old expected `"-0"` was the unverified statement in Univer issue #8270.

## Groups

### 1. `LARGE` ceilings `k`; both functions reject `TRUE`

`fn_large` and `fn_small` in `base/src/functions/statistical/count_and_average.rs`. On 0.9.1 and on `main`, `k` comes from `get_number_no_bools` and is truncated. `1.9` becomes 1, so `LARGE` returns 8. `4.5` becomes 4, which is still inside a four-element range, so it returns 1. `TRUE` is `#VALUE!` because booleans are rejected. Excel's rule, as stated in [#1750](https://github.com/ironcalc/IronCalc/issues/1750), is that `SMALL` truncates and `LARGE` is that rule applied to `count + 1 - k`, which ceilings `k`, and that `TRUE` is 1.

`main` does not fix it. The functions are unchanged since 0.9.1. Open [#1748](https://github.com/ironcalc/IronCalc/pull/1748) calls `fn_large` and `fn_small` for `AGGREGATE` 14 and 15 and does not change those two functions, so it does not cover this. [#1750](https://github.com/ironcalc/IronCalc/issues/1750) is open and has no pull request.

The scratch diff above fixes all four cells. Size is observed: two call-site edits in one file.

### 2. `COUNTA` skips an omitted argument

`fn_counta` in `base/src/functions/statistical/count_and_average.rs`. The parser (`parse_function_args` in `base/src/expressions/parser/mod.rs`) already builds `EmptyArgKind` for `COUNTA(1,)`. `fn_counta` treats `EmptyArg` like `EmptyCell` and does not count it, so the result is 1. Excel counts the omitted argument.

`main` does not fix it. No IronCalc issue or pull request turned up for an omitted `COUNTA` argument. [#1414](https://github.com/ironcalc/IronCalc/issues/1414) is about full-column walks, not this.

The scratch diff counts `EmptyArg` and still skips `EmptyCell`. Proven on the fixture. Size is observed.

### 3. `T.INV(0,1)` is `#DIV/0!`

`fn_t_inv` in `base/src/functions/statistical/t_dist.rs`. Any probability `<= 0` or `>= 1` returns `#NUM!` before the inverse cdf runs. Excel's page says `#NUM!` for probability `<= 0`, and that page is wrong for one degree of freedom: adamlocalis's desktop Excel run on 2026-11-22 got `#DIV/0!` for `T.INV(0,1)` and `#NUM!` for `T.INV(0,2)`. Sources: [fixtures.json](../fixtures/fixtures.json), [HyperFormula's differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences), [Microsoft T.INV](https://support.microsoft.com/en-us/office/t-inv-function-2908272b-4e61-4942-9df9-a25fec9b0e2e).

`main` does not fix it. [#1321](https://github.com/ironcalc/IronCalc/issues/1321) is large arguments. [#1323](https://github.com/ironcalc/IronCalc/issues/1323) is `T.INV.2T` with probability above 1. Neither is this cell, and there is no pull request for it.

The scratch diff returns `#DIV/0!` only when the truncated degrees of freedom are 1 and the probability is 0. Proven. The existing `T.INV(0, 10)` assertion stays `#NUM!`. Size is observed. Treating a non-integer such as `1.9` as 1 is what the current truncation already does; Excel's result for that input was not in the fixture.

### 4. xlsx import intersects both branches of `IF`

`get_function_args_signature` in `base/src/expressions/parser/static_analysis.rs` declares `Function::If` with `args_signature_scalars`, so every argument is a scalar. `add_implicit_intersection`, called on import from `xlsx/src/import/worksheets.rs`, wraps a range in that position with `@`. On `main` the loaded formulas are `SUM(IF(TRUE,@B1:B4))` and `SUM(IF(B1>5,@B1:B4,@C1:C4))`. The implicit intersection misses the formula's row (`E6` and `E7` are below the range) and the sum is `#VALUE!`. Typing the formula in IronCalc does not run that import pass, which is why [#1749](https://github.com/ironcalc/IronCalc/issues/1749) says a typed `SUM(IF(TRUE,B1:B4))` is already 17.

`main` does not fix it. [#1749](https://github.com/ironcalc/IronCalc/issues/1749) is open. No pull request changes the `IF` signature.

The scratch diff marks only the condition as scalar. Proven: the loaded formulas lose `@` and the cells are 17 and 100. The parser test that locked the old `@` on both branches was updated and passed. Size is observed.

### 5. xlsx import does not intersect a `SUMIFS` criterion

Same signature table. `Function::Sumifs` is `vec![Signature::Vector; arg_count]`, so `$A:$A` is left as a whole column. `fn_sumifs` (`base/src/functions/math_and_trigonometry/mathematical.rs`) calls `apply_ifs` (`base/src/functions/statistical/if_ifs.rs`), which hands the criterion to `build_criteria` (`base/src/functions/util.rs`). A `CalcResult::Range` criterion matches nothing, so the sum is 0. Excel shows the legacy formula as `@$A:$A` and sums the same row's key. Citation: [Univer #8154](https://github.com/dream-num/univer/issues/8154) and the implicit-intersection note in fixtures.json.

`main` does not fix it. The loaded formula on `f936620` is still `SUMIFS(Data!B:B,Data!A:A,$A:$A)` and both cells are 0. No IronCalc issue states this fixture. [#1601](https://github.com/ironcalc/IronCalc/issues/1601) is defined names. [#1700](https://github.com/ironcalc/IronCalc/pull/1700) checks that `*IFS` ranges have the same size. [#1753](https://github.com/ironcalc/IronCalc/pull/1753) clips whole-column reads to the used area. Neither adds `@` on the criterion.

The scratch diff marks criteria positions (index 2, 4, ...) as scalar. Proven: both cells match, and `SUMIFS.xlsx`, `SUMIF_AVERAGE_IF.xlsx`, `SUMIF_array.xlsx`, and `COUNTIFS.xlsx` in IronCalc's workbook suite still pass. Size is observed.

### 6. Lookup index below 1 is `#REF!` on 0.9.1

`fn_vlookup` and `fn_hlookup` in `base/src/functions/lookup_and_reference/mod.rs`. On 0.9.1 one test covers both "below 1" and "past the table": `if column_index < 1 || column_index > table.columns()` returns `#REF!` (the same shape in `fn_hlookup`). Excel uses `#VALUE!` when the index is below 1 after truncation, and `#REF!` when it is past the table.

Fixed on `main` by [#1762](https://github.com/ironcalc/IronCalc/pull/1762) (`fc1ddde0`). The run returns `#VALUE!` for both cells. [#1751](https://github.com/ironcalc/IronCalc/issues/1751) is closed. `test_vlookup_column_index_below_one` asserts the split.

### 7. Empty optional arguments in `SORT` and `XMATCH`

`fn_sort` in `base/src/functions/spill_functions.rs` and `fn_xmatch` in `base/src/functions/lookup_and_reference/xmatch.rs`. On 0.9.1 an empty argument is `EmptyArg`, `get_number` turns that into 0, and 0 is an invalid sort index or search mode, so the result is `#VALUE!`. `SORT(A1:A4,,-1)` fails on the empty sort index. `XMATCH(8,A1:A4,,)` fails on the empty search mode (match mode 0 happens to be exact match).

Fixed on `main` by [#1696](https://github.com/ironcalc/IronCalc/pull/1696) (`c4285dda`). The run returns 8 and 3. [#1597](https://github.com/ironcalc/IronCalc/issues/1597) is closed. `fn_xlookup` still has the search-mode half of this bug; that is group 16.

### 8. Date serial 0 is out of range on 0.9.1

`date_part_fn!` in `base/src/functions/date_and_time.rs` expands to `fn_year`, `fn_month`, and `fn_day`. On 0.9.1 it calls `from_excel_date` (`base/src/formatter/dates.rs`), which rejects a serial below `MINIMUM_DATE_SERIAL_NUMBER` (1 in `base/src/constants.rs`) as `#NUM!`. `fn_weekday` calls `from_excel_date` the same way. A blank cell is serial 0, so `YEAR(B1)` fails with it.

Fixed on `main` by [#1758](https://github.com/ironcalc/IronCalc/pull/1758). `serial_to_ymd` (`base/src/formatter/dates.rs`) maps Excel 1900 serial 0 to year 1900, month 1, day 0. `fn_weekday` uses `serial_to_weekday`, and Saturday with return type 1 is 7. The fixture imports as Excel 1900, and the run returns 1900, 0, 7, and 1900. [#1569](https://github.com/ironcalc/IronCalc/issues/1569) is closed. `excel_1900_system_calendar_below_61` asserts `DAY(0)` and `YEAR(0)`. `YEAR(B1)` is the same unary applied to a blank cell, which is serial 0. A new workbook created inside IronCalc still uses the IronCalc date system, where serial 0 is 1899-12-30; that is a different calendar from these xlsx fixtures.

### 9. `AND` returns at the first `FALSE`

`fn_and` in `base/src/functions/logical/and_or_xor_not.rs` calls `logical_nary`. On 0.9.1 that function takes a `short_circuit_value` and returns as soon as `AND` sees `FALSE`, so it never evaluates `A1+B1`. Both inputs are empty text, the sum is `#VALUE!` in Excel, and IronCalc returns `FALSE`.

Fixed on `main` by [#1675](https://github.com/ironcalc/IronCalc/pull/1675) (`8966853a`), which closes [#1609](https://github.com/ironcalc/IronCalc/issues/1609). The run returns `#VALUE!`. `fn_and_or_evaluate_every_argument` covers this shape.

### 10. `DECIMAL` parses a leading sign

`fn_decimal` in `base/src/functions/math_and_trigonometry/mathematical.rs`. On 0.9.1 the text is passed to `i64::from_str_radix`, which accepts `+` and `-`. `DECIMAL("-10",2)` is -2. `DECIMAL(-5,10)` stringifies the number to `-5` and does the same.

Fixed on `main` by [#1660](https://github.com/ironcalc/IronCalc/pull/1660) (`93e11073`), which closes [#1615](https://github.com/ironcalc/IronCalc/issues/1615). A leading `+` or `-` after trimming spaces returns `#NUM!`. The run matches both cells. `decimal_sign_is_num_error` asserts them.

### 11. `SUBTOTAL` 110 and 111, and a product of no numbers

`fn_subtotal` and `subtotal_product` in `base/src/functions/subtotal.rs`. On 0.9.1, function numbers 110 and 111 call the variance helpers with `SubTotalMode::Full`, so a hidden row is included. Row 2 is hidden and `H1:H3` is 2, 8, 4, which is why 0.9.1 returns the three-value variances. `subtotal_product` starts at 1 and returns 1 when the range holds no numbers. Excel's product of no numbers is 0, matching `PRODUCT`.

Fixed on `main` by [#1740](https://github.com/ironcalc/IronCalc/pull/1740) (`30744338`). 110 and 111 use `SkipHidden`. An empty product returns 0. The run returns 0, 2, and 1. [#1739](https://github.com/ironcalc/IronCalc/issues/1739) stays open because of group 12.

### 12. A nested `SUBTOTAL` counts unless it is the whole formula

`cell_is_subtotal` in `base/src/functions/subtotal.rs`. It returns true only when the parsed formula's root node is `Function::Subtotal`. `D6` is `SUBTOTAL(9,D3:D5)+1`, an addition, so `SUBTOTAL(9,D3:D6)` adds that 7 and returns 13. Excel skips any cell whose formula contains `SUBTOTAL`, and the total of 1, 2, and 3 is 6.

`main` does not fix it. The run is still 13. [#1745](https://github.com/ironcalc/IronCalc/pull/1745) is open against [#1739](https://github.com/ironcalc/IronCalc/issues/1739) and was not applied here. Size guess: medium. The walker has to see `SUBTOTAL` anywhere in the tree, and the same helper is used by several `SUBTOTAL` aggregates.

### 13. `TEXT` format `General`

`fn_text` in `base/src/functions/text/common.rs` calls `format_number` in `base/src/formatter/format.rs`. The `ParsePart::General` arm formats with `{value:.9}`, trims trailing zeros, and switches to scientific notation outside `1e-8 .. 1e11` using six significant digits and a signed exponent. Excel's General text, as described in [#1752](https://github.com/ironcalc/IronCalc/issues/1752), keeps at most 11 characters.

`main` does not fix it. The three strings are unchanged. [#1752](https://github.com/ironcalc/IronCalc/issues/1752) is open. No pull request edits that General arm. [#1678](https://github.com/ironcalc/IronCalc/pull/1678) is scientific format codes, not General. Size guess: medium. The arm is local, but the 11-character rule has several shapes (fixed, scientific, and the `1E-09` case that this suite does not store because IronCalc rewrites the literal).

### 14. Text and booleans typed as arguments

Four call sites, one coercion rule, open pull requests, so this was not in the five:

- `text_num_arg` in `base/src/functions/text/common.rs`, used by `fn_left`, rejects a string count. `LEFT("a string","6")` is `#VALUE!`.
- `fn_sumsq` in `base/src/functions/math_and_trigonometry/mathematical.rs` ignores a string argument, so only `4*4` is kept and the result is 16.
- `fn_stdev_s` in `base/src/functions/statistical/standard_dev.rs` ignores a non-number. With no numeric values, count is below 2 and the result is `#DIV/0!`.
- `fn_var_s` in `base/src/functions/statistical/variance.rs` ignores booleans the same way and returns `#DIV/0!`.

`main` does not fix them. The run matches 0.9.1. [#1605](https://github.com/ironcalc/IronCalc/issues/1605) is open. [#1671](https://github.com/ironcalc/IronCalc/pull/1671) covers `LEFT` / `RIGHT` / `MID`. [#1672](https://github.com/ironcalc/IronCalc/pull/1672) covers `STDEV`, `VAR`, and `SUMSQ`. Neither is merged. Size guess: small inside each function once the pull request's coercion is the one Excel uses, and easy to get wrong because a value typed as an argument is counted and the same value read from a cell often is not.

### 15. `XLOOKUP` return array must be the same shape as the lookup vector

`fn_xlookup` in `base/src/functions/xlookup.rs`. After both arguments are ranges, it requires the same row count and the same column count. `A1:A3` is one column and `B1:C3` is two, so it returns `#VALUE!` with "Arrays must be of the same size". Excel returns the matching row and spills. The fixture scores the first spill cell, `x`. adamlocalis entered this as a dynamic-array formula on 2026-11-23, with E1 cleared, and it spilled `x` and `y`. The legacy non-array form of this formula returns `#VALUE!` in Excel (implicit intersection), and the dynamic-array form returns `x`.

`main` does not fix it. [#1789](https://github.com/ironcalc/IronCalc/issues/1789) is open. [#1703](https://github.com/ironcalc/IronCalc/pull/1703) and [#1440](https://github.com/ironcalc/IronCalc/pull/1440) are open and describe multi-column returns. They were not applied. Size guess: medium. The match has to return a row, and the cell has to spill.

### 16. Empty `XLOOKUP` search mode becomes 0

Same function. `get_number` maps `EmptyArg` to 0 (`cast_to_number` in `base/src/cast.rs`). Match mode 0 is exact, so the fifth argument being empty is harmless. Search mode 0 is not 1, -1, 2, or -2, and `fn_xlookup` returns `#ERROR!`. Excel defaults an empty search mode to 1 (first to last). `fn_xmatch` on `main` already special-cases `EmptyArgKind` for this.

`main` does not fix it. The loaded formula is still `XLOOKUP(8,H1:H4,I1:I4,,,)` and the value is `#ERROR!`. [#1718](https://github.com/ironcalc/IronCalc/issues/1718) is open. [#1721](https://github.com/ironcalc/IronCalc/pull/1721) is the open pull request and was not applied. Size guess: small, the same `EmptyArgKind` check `fn_xmatch` already has.

### 17. `NPV` ignores a boolean written in the formula

`fn_npv` in `base/src/functions/financial/mod.rs`. A direct argument that evaluates to a boolean falls through to the arm that ignores non-numbers. `TRUE()` is dropped, the remaining `1` is the first cash flow, and `1/2` is 0.5. Excel counts a boolean written in the formula as 1 or 0 and keeps its position, so `TRUE()` and then `1` is `1/2 + 1/4 = 0.75`. A boolean stored in a cell is ignored by both: `NPV(1,B1)` with B1 TRUE is 0 and is not a miss.

`main` does not fix it. [#1594](https://github.com/ironcalc/IronCalc/issues/1594) is `NPV` ignoring array arguments, not a boolean literal. No pull request covers `TRUE()`. Size guess: small. The node has to be classified before evaluation, because a cell reference and `TRUE()` both arrive as `CalcResult::Boolean`.

### 18. Weibull pdf at 0

`fn_weibull_dist` in `base/src/functions/statistical/weibull.rs`. For `x = 0`, `alpha = 1`, `beta = 1`, cumulative false, it uses `statrs`' pdf. That pdf is `(1/beta) * exp(0) = 1`. Excel returns 0. Citation: adamlocalis, 2026-11-22, plus [HyperFormula](https://hyperformula.handsontable.com/docs/guide/list-of-differences) and [Microsoft Q&A 5065194](https://learn.microsoft.com/en-us/answers/questions/5065194/weibull-dist-function-error).

`main` does not fix it. The run is 1. No IronCalc issue or pull request names `WEIBULL.DIST`. [#55](https://github.com/ironcalc/IronCalc/issues/55) and [#774](https://github.com/ironcalc/IronCalc/issues/774) are the general statistical-function lists. Size guess: small for this one input. Returning 0 for every non-cumulative `x = 0` may be wrong when `alpha < 1`, and that input is not in the suite.

### 19. `TEXT` of a negative number that rounds to zero

Not a miss. `fn_text` calls `format_number`. Format `"0"` is one section. In the digit loop, `is_negative` is `value < -(10^(-precision))`. Precision 0 makes the threshold -1, so -0.4 is formatted as `"0"`. That matches Excel. The old expected `"-0"` came from Univer [#8270](https://github.com/dream-num/univer/issues/8270), whose author says they did not run it in Excel. adamlocalis's desktop Excel (16.0300) run on 2026-11-23 returned `"0"`. The MrExcel thread and the Microsoft Q&A linked from the fixture were inferred support that Excel chooses the negative section from the sign before rounding; that guess was wrong for this formula.

`main` also returns `"0"`. [#1754](https://github.com/ironcalc/IronCalc/pull/1754) and [#1697](https://github.com/ironcalc/IronCalc/pull/1697) both replace `is_negative` and still require the rounded digits to be non-zero before they print a minus, so a rounded zero stays unsigned, which is what Excel returns here.

### 20. Accounting parentheses are not numbers

`fn_value` in `base/src/functions/text/common.rs` parses with `parse_formatted_number` (`base/src/formatter/format.rs`). `"(5)"+1` reaches the same parser through `cast_number` (`base/src/cast.rs`). The parser understands a leading currency sign and a trailing percent sign. It does not understand `(5)` or `($1,000)` as negatives, so both `VALUE` and the addition return `#VALUE!`. Citation: [Univer #8231](https://github.com/dream-num/univer/issues/8231).

`main` does not fix it. All three cells are `#VALUE!`. No IronCalc issue turned up. [#1443](https://github.com/ironcalc/IronCalc/pull/1443) teaches `VALUE` and `DATEVALUE` to read dates and times written as text; its diff does not mention parentheses. Size guess: small to medium, one parser used by `VALUE` and by arithmetic, with currency and thousands separators inside the parentheses.

## What `main` already matches

Fourteen of the thirty-nine 0.9.1 misses match Excel on `f936620`: groups 6, 7, 8, 9, 10, and 11 (F15, F19, F20, F22, F25, and F24 `A1`/`A8`/`A9`). The other 25 are still misses. The five scratch edits cover ten of those 25. The remaining 15 are groups 12 through 18 and group 20. Seven of those 15 sit under an open pull request (group 12, group 14, groups 15 and 16) and were left alone.
