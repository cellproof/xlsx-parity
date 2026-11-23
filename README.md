# xlsx-parity

Excel formula parity and XLSX round-trip checks for open spreadsheet libraries.

`make test` reproduces the baseline (25 fixtures, 73 formula cells, plus `fixtures/RT-features.xlsx`) on IronCalc's Python bindings and on IronCalc's JavaScript build. Expected values are Excel's. Each one is copied from a Univer issue, an IronCalc issue, [HyperFormula's list of differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences), or a Microsoft Support page, and the citation is in [`fixtures/fixtures.json`](fixtures/fixtures.json). The suite does not invent expected values. Milestone 4's new cases are listed in [`CHANGELOG.md`](CHANGELOG.md).

## Run

```sh
make test
```

Needs Python 3.10+ with the `venv` module, Node.js, and `make`. Network is needed the first time. The command creates `.venv`, installs the pinned Python packages, installs SheetJS Community Edition 0.20.3 from its CDN tarball and `@ironcalc/nodejs` 0.9.1, and runs the suite.

Exit 0 means the measured results match [`suite/baseline.json`](suite/baseline.json), including the known IronCalc failures. A regression or an unexpected fix exits nonzero. The JavaScript build has its own recorded results in that file. A change in a JavaScript result fails the run the same way.

## Baseline

IronCalc 0.9.1 (Python bindings) and `@ironcalc/nodejs` 0.9.1 calculate. SheetJS CE 0.20.3 does not; it is an I/O witness for formula strings and cached values.

- IronCalc Python calc parity: 33/73 cells pass (33/72 excluding F04), 10 of 25 fixtures fully pass, 0 errors.
- IronCalc JS (`@ironcalc/nodejs` 0.9.1) calc parity: 33/73 cells match Excel, 73/73 match the Python bindings, 0 errors. This package loads the xlsx fixtures with `Model.fromXlsx`. `@ironcalc/wasm` is the engine without an xlsx reader, so it is not used.
- Expected IronCalc failures, same on Python and JS, are the 40 cells in `ironcalc_failures`. The original three are F10 `SUMIFS` with whole-column criteria returns 0 (Excel: 10809 and 10800), F07 `VALUE` of accounting negatives like `"(5)"` returns `#VALUE!` (Excel: -5, -1000, -4), and F04 `TEXT(-0.4,"0")` returns `"0"` (Excel: `"-0"`). F04 is the weaker citation: the issue author says they could not check it in Excel. The other 34 are the Milestone 4 misses in [`CHANGELOG.md`](CHANGELOG.md).
- SheetJS CE read: 73/73 formula strings and cached values preserved. The write check is the same 73/73.
- IronCalc keeps the 73 formula strings and writes its 40 wrong values back as cached values. That set is an assertion, `ironcalc_wrong_cached_cells`. A cell joining or leaving the set fails the run and names the cell.
- On `RT-features.xlsx`, SheetJS CE loses bold/fill, freeze panes, conditional formatting, data validation, and the chart. IronCalc loses the comment, data validation, and the chart. Each feature is read from the written workbook for each library. The element is named on the check in [`suite/run.py`](suite/run.py). A feature that is newly kept or newly lost fails the run and names the feature and the library.

## Layout

- `fixtures/` — committed workbooks, `fixtures.json`, and the two scripts that built them (`build_fixtures.py`, `build_features.py`). `make test` does not regenerate the workbooks.
- `suite/run.py` — the check.
- `suite/ironcalc_js.js` — calculates the same cells with `@ironcalc/nodejs`.
- `suite/baseline.json` — recorded IronCalc Python failures, the wrong cached cells, IronCalc JS results, and the feature table.
- `CHANGELOG.md` — each Milestone 4 case, its Excel value, source, and IronCalc's result.

## Licenses

This project is [MIT](LICENSE).

Installed by `make test`, not vendored:

| Package | Version | License |
|---|---|---|
| [ironcalc](https://github.com/ironcalc/IronCalc) (Python bindings) | 0.9.1 | MIT OR Apache-2.0 |
| [@ironcalc/nodejs](https://www.npmjs.com/package/@ironcalc/nodejs) | 0.9.1 | MIT OR Apache-2.0 |
| [SheetJS Community Edition](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz) (`xlsx`) | 0.20.3 | Apache-2.0 |
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.5 | MIT |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) | 3.3.4 | BSD-2-Clause |

SheetJS Community Edition only. This project does not use SheetJS Pro.

`@ironcalc/nodejs` 0.9.1 ships both `LICENSE-MIT` and `LICENSE-Apache-2.0` (SPDX `MIT OR Apache-2.0`). npm also installs the matching platform binary. Those binaries declare MIT; on linux x64 this is `@ironcalc/nodejs-linux-x64-gnu` 0.9.1. No paid or non-open-source packages are installed.

## Limits

Expected values were taken from issue text, HyperFormula's published Excel column, and Microsoft docs, not from running Excel. Seventy-three cases is a small sample, not a draw of Excel as a whole. Univer and LibreOffice are out of this run.
