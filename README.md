# xlsx-parity

Excel formula parity and XLSX round-trip checks for open spreadsheet libraries.

`make test` reproduces the baseline (25 fixtures, 73 formula cells, plus `fixtures/RT-features.xlsx`) on IronCalc's Python bindings, on IronCalc's JavaScript build, and on LibreOffice Calc. Expected values are Excel's. Each one is copied from a Univer issue, an IronCalc issue, [HyperFormula's list of differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences), or a Microsoft Support page, and the citation is in [`fixtures/fixtures.json`](fixtures/fixtures.json). The suite does not invent expected values. Milestone 4's new cases are listed in [`CHANGELOG.md`](CHANGELOG.md).

## Run

```sh
make test
```

Needs Python 3.10+ with the `venv` module, Node.js, `make`, and LibreOffice Calc `soffice` on `PATH`. Network is needed the first time. The command creates `.venv`, installs the pinned Python packages, installs SheetJS Community Edition 0.20.3 from its CDN tarball, `@ironcalc/nodejs` 0.9.1, and Univer 1.6.0 (`@univerjs/core`, `@univerjs/docs`, `@univerjs/engine-formula`, `@univerjs/engine-render`, `@univerjs/sheets`, `@univerjs/sheets-formula`, `@univerjs/sheets-numfmt`, and `rxjs` 7.8.2), and runs the suite.

LibreOffice is not installed by `make test`. CI installs the pinned Ubuntu 24.04 backport (the runner is `ubuntu-24.04`, because `ubuntu-latest` is 26.04 and that archive's Calc package is a different revision):

```sh
sudo apt-get update
sudo apt-get install -y -t noble-backports --no-install-recommends \
  libreoffice-calc=4:26.2.5.2-0ubuntu0.26.04.1~bpo24.04.1
```

Exit 0 means the measured results match [`suite/baseline.json`](suite/baseline.json), including the known IronCalc failures, Univer's recorded native results, the engine-hook results on all 73 cells, and LibreOffice's recorded results. A regression or an unexpected fix exits nonzero. The JavaScript build has its own recorded results in that file. A change in a JavaScript result fails the run the same way. A change in a Univer result fails the run the same way. A change in a LibreOffice result fails the run the same way. A LibreOffice miss that matches the record is a finding.

## Baseline

IronCalc 0.9.1 (Python bindings) and `@ironcalc/nodejs` 0.9.1 calculate. SheetJS CE 0.20.3 does not; it is an I/O witness for formula strings and cached values.

- IronCalc Python calc parity: 33/73 cells pass (33/72 excluding F04), 10 of 25 fixtures fully pass, 0 errors.
- IronCalc JS (`@ironcalc/nodejs` 0.9.1) calc parity: 33/73 cells match Excel, 73/73 match the Python bindings, 0 errors. This package loads the xlsx fixtures with `Model.fromXlsx`. `@ironcalc/wasm` is the engine without an xlsx reader, so it is not used.
- Univer 1.6.0 headless, open-source packages only. SheetJS CE reads each formula workbook into an `IWorkbookData` snapshot (`suite/univer_adapter.js`). Univer's own formula engine is scored on all 25 formula fixtures (73 cells): 45/73 match Excel. The misses are findings and are locked in the baseline. A second pass subclasses `CalculateFormulaService` and binds it with `UniverFormulaEnginePlugin._initializeWithOverride`, so `@ironcalc/nodejs` calculates all 73 cells through `Model.fromXlsx`. The hook matches that direct IronCalc output on every cell (73/73) and matches Excel on 33/73. At startup the override checks that `_initializeWithOverride` and the `ICalculateFormulaService` binding are still present. If either is missing, the run stops with one error naming the Univer version: the engine-swap hook moved. `@univerjs/preset-sheets-core` is a browser preset, so this run uses plugin mode.
- LibreOffice Calc 26.2.5.2 (`soffice --version` prints `LibreOffice 26.2.5.2 620(Build:2)`; Debian package `libreoffice-calc` `4:26.2.5.2-0ubuntu0.26.04.1~bpo24.04.1`). Headless `soffice --convert-to xlsx` keeps stored formula caches unless the user profile sets `OOXMLRecalcMode` to 0. The suite deletes those caches in a temporary copy, converts all 25 formula workbooks in one process, and reads the new cached values from the xlsx. 45/73 cells match Excel. The misses are findings and are locked in the baseline. `make test` prints each cell, then a per-fixture count for IronCalc, Univer native, and LibreOffice. F04 `Sheet1!A1` is the recalc proof: the file's stored cache is `-0`, a convert without the setting still writes `-0`, and the cleared-cache convert writes `0`. A change in that triple fails the run. The count of cells whose LibreOffice value differs from the stored cache is printed; zero fails the run.
- Expected IronCalc failures, same on Python and JS, are the 40 cells in `ironcalc_failures`. The original three are F10 `SUMIFS` with whole-column criteria returns 0 (Excel: 10809 and 10800), F07 `VALUE` of accounting negatives like `"(5)"` returns `#VALUE!` (Excel: -5, -1000, -4), and F04 `TEXT(-0.4,"0")` returns `"0"` (Excel: `"-0"`). F04 is the weaker citation: the issue author says they could not check it in Excel. The other 34 are the Milestone 4 misses in [`CHANGELOG.md`](CHANGELOG.md).
- SheetJS CE read: 73/73 formula strings and cached values preserved. The write check is the same 73/73.
- IronCalc keeps the 73 formula strings and writes its 40 wrong values back as cached values. That set is an assertion, `ironcalc_wrong_cached_cells`. A cell joining or leaving the set fails the run and names the cell.
- On `RT-features.xlsx`, SheetJS CE loses bold/fill, freeze panes, conditional formatting, data validation, and the chart. IronCalc loses the comment, data validation, and the chart. Each feature is read from the written workbook for each library. The element is named on the check in [`suite/run.py`](suite/run.py). A feature that is newly kept or newly lost fails the run and names the feature and the library.

## Layout

- `fixtures/` — committed workbooks, `fixtures.json`, and the two scripts that built them (`build_fixtures.py`, `build_features.py`). `make test` does not regenerate the workbooks.
- `suite/run.py` — the check.
- `suite/ironcalc_js.js` — calculates the same cells with `@ironcalc/nodejs`.
- `suite/univer_adapter.js` — loads formula workbooks into headless Univer, scores Univer's engine, and runs the IronCalc engine hook on every formula cell.
- `suite/libreoffice_calc.py` — recalculates the formula fixtures with headless LibreOffice Calc and reads the new cached values.
- `suite/baseline.json` — recorded IronCalc Python failures, the wrong cached cells, IronCalc JS results, Univer native results, the engine-hook results, LibreOffice results, and the feature table.
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
| [@univerjs/core](https://www.npmjs.com/package/@univerjs/core) | 1.6.0 | Apache-2.0 |
| [@univerjs/docs](https://www.npmjs.com/package/@univerjs/docs) | 1.6.0 | Apache-2.0 |
| [@univerjs/engine-formula](https://www.npmjs.com/package/@univerjs/engine-formula) | 1.6.0 | Apache-2.0 |
| [@univerjs/engine-render](https://www.npmjs.com/package/@univerjs/engine-render) | 1.6.0 | Apache-2.0 |
| [@univerjs/sheets](https://www.npmjs.com/package/@univerjs/sheets) | 1.6.0 | Apache-2.0 |
| [@univerjs/sheets-formula](https://www.npmjs.com/package/@univerjs/sheets-formula) | 1.6.0 | Apache-2.0 |
| [@univerjs/sheets-numfmt](https://www.npmjs.com/package/@univerjs/sheets-numfmt) | 1.6.0 | Apache-2.0 |
| [rxjs](https://github.com/ReactiveX/rxjs) | 7.8.2 | Apache-2.0 |

Installed in CI from Ubuntu, not by `make test` and not vendored:

| Package | Version | License |
|---|---|---|
| [LibreOffice Calc](https://www.libreoffice.org/) (`libreoffice-calc`) | 26.2.5.2 | MPL-2.0 |

SheetJS Community Edition only. This project does not use SheetJS Pro.

Univer open-source packages only, release [v1.6.0](https://github.com/dream-num/univer/releases/tag/v1.6.0) (`eaf4cfa04fabdbcfc922c692f3450459f8a29546`). This project does not use `@univerjs-pro` or Univer Server. `npm ci` also installs the dependencies those packages declare. Each of those packages declares an OSI license (Apache-2.0, MIT, BSD-3-Clause, ISC, or 0BSD).

`@ironcalc/nodejs` 0.9.1 ships both `LICENSE-MIT` and `LICENSE-Apache-2.0` (SPDX `MIT OR Apache-2.0`). npm also installs the matching platform binary. Those binaries declare MIT; on linux x64 this is `@ironcalc/nodejs-linux-x64-gnu` 0.9.1. No paid or non-open-source packages are installed.

LibreOffice Calc is installed from Ubuntu's `libreoffice-calc` package, not vendored. The package copyright is MPL-2.0, and some files in that upstream tarball are Apache-2.0. This project does not install a LibreOffice language pack.

## Limits

Expected values were taken from issue text, HyperFormula's published Excel column, and Microsoft docs, not from running Excel. Seventy-three cases is a small sample, not a draw of Excel as a whole. Univer's own engine, the IronCalc engine hook, and LibreOffice Calc are scored on those 73 cells. LibreOffice misses that match the baseline are findings.
