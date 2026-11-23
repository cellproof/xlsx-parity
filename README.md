# xlsx-parity

Excel formula parity and XLSX round-trip checks for open spreadsheet libraries.

`make test` reproduces the Test 1 baseline (10 fixtures, 27 formula cells, plus `fixtures/RT-features.xlsx`) on IronCalc's Python bindings and on IronCalc's JavaScript build. Expected values are Excel's. Each one is copied from a Univer issue or a Microsoft Support page, and the citation is in [`fixtures/fixtures.json`](fixtures/fixtures.json). The suite does not invent expected values.

## Run

```sh
make test
```

Needs Python 3.10+ with the `venv` module, Node.js, and `make`. Network is needed the first time. The command creates `.venv`, installs the pinned Python packages, installs SheetJS Community Edition 0.20.3 from its CDN tarball and `@ironcalc/nodejs` 0.9.1, and runs the suite.

Exit 0 means the measured results match [`suite/baseline.json`](suite/baseline.json), including the known IronCalc failures. A regression or an unexpected fix exits nonzero. The JavaScript build has its own recorded results in that file. A change in a JavaScript result fails the run the same way.

## Baseline

IronCalc 0.9.1 (Python bindings) and `@ironcalc/nodejs` 0.9.1 calculate. SheetJS CE 0.20.3 does not; it is an I/O witness for formula strings and cached values.

- IronCalc Python calc parity: 21/27 cells pass (21/26 excluding F04), 7 of 10 fixtures fully pass, 0 errors.
- IronCalc JS (`@ironcalc/nodejs` 0.9.1) calc parity: 21/27 cells match Excel, 27/27 match the Python bindings, 0 errors. This package loads the xlsx fixtures with `Model.fromXlsx`. `@ironcalc/wasm` is the engine without an xlsx reader, so it is not used.
- Expected IronCalc failures, same on Python and JS: F10 `SUMIFS` with whole-column criteria returns 0 (Excel: 10809 and 10800), F07 `VALUE` of accounting negatives like `"(5)"` returns `#VALUE!` (Excel: -5, -1000, -4), F04 `TEXT(-0.4,"0")` returns `"0"` (Excel: `"-0"`). F04 is the weaker citation: the issue author says they could not check it in Excel.
- SheetJS CE read: 27/27 formula strings and cached values preserved. The write check is the same 27/27.
- IronCalc keeps the 27 formula strings and writes its 6 wrong values back as cached values.
- On `RT-features.xlsx`, SheetJS CE loses bold/fill, freeze panes, conditional formatting, data validation, and the chart. IronCalc loses the comment, data validation, and the chart.

## Layout

- `fixtures/` — committed workbooks, `fixtures.json`, and the two scripts that built them (`build_fixtures.py`, `build_features.py`). `make test` does not regenerate the workbooks.
- `suite/run.py` — the check.
- `suite/ironcalc_js.js` — calculates the same cells with `@ironcalc/nodejs`.
- `suite/baseline.json` — recorded IronCalc Python failures, IronCalc JS results, and the feature table.

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

Expected values were taken from issue text and Microsoft docs, not from running Excel. Twenty-seven cases is a small sample of bugs Univer had, not a draw of Excel as a whole. Univer and LibreOffice are out of this run.
