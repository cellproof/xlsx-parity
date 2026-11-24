# xlsx-parity

This suite measures formula parity and XLSX round-trip behavior for open spreadsheet libraries. It scores 25 fixtures (74 formula cells). Expected values come from recorded desktop Excel runs, plus the remaining cited sources. It also checks which features of `fixtures/RT-features.xlsx` survive a write. Scores are in [SCOREBOARD.md](SCOREBOARD.md).

Run it with `make test`.

`make test` on this tree matched [`suite/baseline.json`](suite/baseline.json) (`baseline: match`). The summary was:

| Check | Result |
|---|---|
| IronCalc 0.9.1 (Python bindings) | 35/74 cells match Excel |
| Univer 1.6.0 engine hook (`@ironcalc/nodejs` 0.9.1) | 35/74 cells match Excel |
| Univer 1.6.0 native (`@univerjs/engine-formula`) | 49/74 cells match Excel |
| LibreOffice Calc 26.2.5.2 | 47/74 cells match Excel |
| SheetJS CE 0.20.3 | 74/74 read and 74/74 write |

`@ironcalc/nodejs` 0.9.1 also matches Excel on 35/74 and matches the Python bindings on 74/74. The Univer hook matches that IronCalc output on 74/74.

A finding is a miss already recorded in `suite/baseline.json`. The run passes when every measurement matches that file, including those misses. A new miss, or a recorded miss that starts matching Excel, fails the run.

Expected values come from recorded desktop Excel runs, plus the remaining cited sources in [`fixtures/fixtures.json`](fixtures/fixtures.json). adamlocalis's desktop Excel (16.0300) run on 2026-11-23 is the primary source for every formula cell except F11–F13, which come from his desktop Excel for Microsoft 365 run on 2026-11-22. Each cell keeps its earlier citation after `previously:`. [HyperFormula's list of differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) is a reference for the F11–F13 formulas, not the source of the values. The suite does not invent expected values. Milestone 4's new cases are listed in [`CHANGELOG.md`](CHANGELOG.md). Numeric comparisons with Excel use a relative tolerance of 1e-12. Strings, booleans, and errors are exact.

The fixture files are written with xlsxwriter, with each expected value stored as the cached value, so SheetJS's score measures round-tripping those values, not Excel-saved values.

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

Exit 0 means the measured results match [`suite/baseline.json`](suite/baseline.json), including the known IronCalc failures, Univer's recorded native results, the engine-hook results on all 74 cells, and LibreOffice's recorded results. A regression or an unexpected fix exits nonzero. The JavaScript build has its own recorded results in that file. A change in a JavaScript result fails the run the same way. A change in a Univer result fails the run the same way. A change in a LibreOffice result fails the run the same way. A LibreOffice miss that matches the record is a finding.

## Baseline

IronCalc 0.9.1 (Python bindings) and `@ironcalc/nodejs` 0.9.1 calculate. SheetJS CE 0.20.3 does not; it is an I/O witness for formula strings and cached values.

- IronCalc Python calc parity: 35/74 cells pass (34/73 excluding F04), 11 of 25 fixtures fully pass, 0 errors.
- IronCalc JS (`@ironcalc/nodejs` 0.9.1) calc parity: 35/74 cells match Excel, 74/74 match the Python bindings, 0 errors. This package loads the xlsx fixtures with `Model.fromXlsx`. `@ironcalc/wasm` is the engine without an xlsx reader, so it is not used.
- Univer 1.6.0 headless, open-source packages only. SheetJS CE reads each formula workbook into an `IWorkbookData` snapshot (`suite/univer_adapter.js`). Univer's own formula engine is scored on all 25 formula fixtures (74 cells): 49/74 match Excel. The misses are findings and are locked in the baseline. One match is F17 `STDEV("1","2")`: Univer returns 0.707106781187, within 1e-12 relative of Excel's 0.7071067811865476. A second pass subclasses `CalculateFormulaService` and binds it with `UniverFormulaEnginePlugin._initializeWithOverride`, so `@ironcalc/nodejs` calculates all 74 cells through `Model.fromXlsx`. The hook matches that direct IronCalc output on every cell (74/74) and matches Excel on 35/74. At startup the override checks that `_initializeWithOverride` and the `ICalculateFormulaService` binding are still present. If either is missing, the run stops with one error naming the Univer version: the engine-swap hook moved. `@univerjs/preset-sheets-core` is a browser preset, so this run uses plugin mode.
- LibreOffice Calc 26.2.5.2 (`soffice --version` prints `LibreOffice 26.2.5.2 620(Build:2)`; Debian package `libreoffice-calc` `4:26.2.5.2-0ubuntu0.26.04.1~bpo24.04.1`). Headless `soffice --convert-to xlsx` keeps stored formula caches unless the user profile sets `OOXMLRecalcMode` to 0. The suite deletes those caches in a temporary copy, converts all 25 formula workbooks in one process, and reads the new cached values from the xlsx. 47/74 cells match Excel. The misses are findings and are locked in the baseline. One match is F17 `VAR(TRUE,FALSE,TRUE)`: LibreOffice returns 0.333333333333333, within 1e-12 relative of Excel's 0.33333333333333337. F12 `T.INV(0,2)` is a locked miss (`#VALUE!`; Excel is `#NUM!`). `make test` prints each cell, then a per-fixture count for IronCalc, Univer native, and LibreOffice. F04 `Sheet1!A1` is the recalc proof: the file's stored cache is `0`, a convert without the setting still writes `0`, and the cleared-cache convert writes `0`. A change in that triple fails the run. The count of cells whose LibreOffice value differs from the stored cache is printed; zero fails the run.
- Expected IronCalc failures, same on Python and JS, are the 39 cells in `ironcalc_failures`. Two of the original findings remain: F10 `SUMIFS` with whole-column criteria returns 0 (Excel: 10809 and 10800), and F07 `VALUE` of accounting negatives like `"(5)"` returns `#VALUE!` (Excel: -5, -1000, -4). F04 `TEXT(-0.4,"0")` returns `"0"`, which matches Excel. The old expected `"-0"` was Univer issue #8270, and the issue author had not checked it in Excel. The other 34 are the Milestone 4 misses in [`CHANGELOG.md`](CHANGELOG.md). F12 `T.INV(0,2)` matches Excel (`#NUM!`) on IronCalc and is not in that failure list.
- SheetJS CE read: 74/74 formula strings and cached values preserved. The write check is the same 74/74.
- IronCalc keeps the 74 formula strings and writes its 39 wrong values back as cached values. That set is an assertion, `ironcalc_wrong_cached_cells`. A cell joining or leaving the set fails the run and names the cell.
- On `RT-features.xlsx`, SheetJS CE loses bold/fill, freeze panes, conditional formatting, data validation, and the chart. IronCalc loses the comment, data validation, and the chart. Each feature is read from the written workbook for each library. The element is named on the check in [`suite/run.py`](suite/run.py). A feature that is newly kept or newly lost fails the run and names the feature and the library.

## Layout

- `fixtures/` — committed workbooks, `fixtures.json`, and the two scripts that built them (`build_fixtures.py`, `build_features.py`). `make test` does not regenerate the workbooks.
- `suite/run.py` — the check.
- `suite/ironcalc_js.js` — calculates the same cells with `@ironcalc/nodejs`.
- `suite/univer_adapter.js` — loads formula workbooks into headless Univer, scores Univer's engine, and runs the IronCalc engine hook on every formula cell.
- `suite/libreoffice_calc.py` — recalculates the formula fixtures with headless LibreOffice Calc and reads the new cached values.
- `suite/baseline.json` — recorded IronCalc Python failures, the wrong cached cells, IronCalc JS results, Univer native results, the engine-hook results, LibreOffice results, and the feature table.
- `CHANGELOG.md` — each Milestone 4 case, its Excel value, source, and IronCalc's result.

## License

Licensed under either of the following, at your option:

- [MIT License](LICENSE-MIT)
- [Apache License, Version 2.0](LICENSE-APACHE)

SPDX identifier: `MIT OR Apache-2.0`.

Copyright (c) 2026 CellProof contributors

Installed packages, the LibreOffice package used in CI, and the fixture sources are listed in [THIRD_PARTY.md](THIRD_PARTY.md). SheetJS Community Edition only. Univer open-source packages only. No SheetJS Pro, `@univerjs-pro`, or other paid package is installed.

## Limits

Expected values come from recorded desktop Excel runs, plus the remaining cited sources. F11–F13, including `T.INV(0,1)` and `T.INV(0,2)`, come from a desktop Excel for Microsoft 365 run by adamlocalis on 2026-11-22. The other formula cells come from his desktop Excel (16.0300) run on 2026-11-23. HyperFormula's differences page is a reference for the F11–F13 formulas only. Seventy-four cases is a small sample, not a draw of Excel as a whole. Univer's own engine, the IronCalc engine hook, and LibreOffice Calc are scored on those 74 cells. LibreOffice misses that match the baseline are findings. Numeric Excel checks use a relative tolerance of 1e-12; strings, booleans, and errors stay exact.
