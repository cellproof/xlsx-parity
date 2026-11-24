# Third-party notices

Nothing in this repository is vendored. `make test` installs the Python and npm packages below. CI installs LibreOffice Calc from Ubuntu and does not commit it. Versions and licenses were read from the lockfile, from each direct package's own license file after `npm ci` and `pip install`, and from the Debian copyright file for `libreoffice-calc`.

This project is dual-licensed `MIT OR Apache-2.0` (see `LICENSE-MIT` and `LICENSE-APACHE`). Copyright (c) 2026 CellProof contributors.

## Direct packages

| Package | Version | License | Where it comes from |
|---|---|---|---|
| ironcalc (Python bindings) | 0.9.1 | MIT OR Apache-2.0 | PyPI, `requirements.txt`. Wheel ships `LICENSE-MIT.md` and `LICENSE-Apache-2.0.md`. Classifiers: MIT and Apache Software License. |
| @ironcalc/nodejs | 0.9.1 | MIT OR Apache-2.0 | npm. Ships `LICENSE-MIT` and `LICENSE-Apache-2.0`. |
| @ironcalc/nodejs-linux-x64-gnu | 0.9.1 | MIT | Optional platform binary. This is the one `npm ci` installs on linux x64. |
| xlsx (SheetJS Community Edition) | 0.20.3 | Apache-2.0 | https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz. Installed `package.json` says `Apache-2.0`. The package README calls this the Community Edition. |
| @univerjs/core | 1.6.0 | Apache-2.0 | npm. Open-source Univer release v1.6.0 (`eaf4cfa04fabdbcfc922c692f3450459f8a29546`). |
| @univerjs/docs | 1.6.0 | Apache-2.0 | npm. |
| @univerjs/engine-formula | 1.6.0 | Apache-2.0 | npm. |
| @univerjs/engine-render | 1.6.0 | Apache-2.0 | npm. |
| @univerjs/sheets | 1.6.0 | Apache-2.0 | npm. |
| @univerjs/sheets-formula | 1.6.0 | Apache-2.0 | npm. |
| @univerjs/sheets-numfmt | 1.6.0 | Apache-2.0 | npm. |
| rxjs | 7.8.2 | Apache-2.0 | npm. Pinned directly. Also required by the Univer packages. |
| openpyxl | 3.1.5 | MIT | PyPI, `requirements.txt`. `__license__` is MIT. Copyright (c) 2010-2024 openpyxl. |
| et-xmlfile | 2.0.0 | MIT | Installed because openpyxl depends on it. Not pinned in `requirements.txt`. `__license__` is MIT. |
| XlsxWriter | 3.3.4 | BSD-2-Clause | PyPI, `requirements.txt`. Used by `fixtures/build_fixtures.py` and `fixtures/build_features.py`. `LICENSE.txt` is BSD-2-Clause, copyright (c) 2013-2026 John McNamara. |
| libreoffice-calc | 4:26.2.5.2-0ubuntu0.26.04.1~bpo24.04.1 | MPL-2.0 | Ubuntu noble-backports, installed in CI only. Debian copyright: MPL-2.0, and some upstream files are Apache-2.0. Not vendored. No language pack is installed. |

Other `@ironcalc/nodejs` 0.9.1 platform packages are optional and declare MIT in the lockfile. `npm ci` on linux x64 does not install them: `@ironcalc/nodejs-android-arm-eabi`, `@ironcalc/nodejs-android-arm64`, `@ironcalc/nodejs-darwin-arm64`, `@ironcalc/nodejs-darwin-universal`, `@ironcalc/nodejs-darwin-x64`, `@ironcalc/nodejs-linux-arm-gnueabihf`, `@ironcalc/nodejs-linux-arm64-gnu`, `@ironcalc/nodejs-linux-riscv64-gnu`, `@ironcalc/nodejs-win32-arm64-msvc`, `@ironcalc/nodejs-win32-x64-msvc`.

`@ironcalc/nodejs` also names `@ironcalc/nodejs-linux-arm-musleabihf`, `@ironcalc/nodejs-linux-arm64-musl`, and `@ironcalc/nodejs-linux-x64-musl` at 0.9.1. Those three names are not separate entries in `package-lock.json`, they are not on the npm registry at 0.9.1, and `npm ci` does not install them. There is no artifact here to license.

## npm packages `npm ci` can install

Every package recorded in `package-lock.json`, including optional platform binaries. Each one has a `license` field. The set is Apache-2.0, MIT, `MIT OR Apache-2.0`, BSD-3-Clause, ISC, and 0BSD. None is missing, and none is a commercial or proprietary license.

| Package | Version | License |
|---|---|---|
| `@flatten-js/interval-tree` | 1.1.3 | MIT |
| `@floating-ui/core` | 1.8.0 | MIT |
| `@floating-ui/dom` | 1.8.0 | MIT |
| `@floating-ui/utils` | 0.2.12 | MIT |
| `@grpc/grpc-js` | 1.14.7 | Apache-2.0 |
| `@grpc/proto-loader` | 0.8.1 | Apache-2.0 |
| `@ironcalc/nodejs` | 0.9.1 | MIT OR Apache-2.0 |
| `@ironcalc/nodejs-android-arm-eabi` | 0.9.1 | MIT |
| `@ironcalc/nodejs-android-arm64` | 0.9.1 | MIT |
| `@ironcalc/nodejs-darwin-arm64` | 0.9.1 | MIT |
| `@ironcalc/nodejs-darwin-universal` | 0.9.1 | MIT |
| `@ironcalc/nodejs-darwin-x64` | 0.9.1 | MIT |
| `@ironcalc/nodejs-linux-arm-gnueabihf` | 0.9.1 | MIT |
| `@ironcalc/nodejs-linux-arm64-gnu` | 0.9.1 | MIT |
| `@ironcalc/nodejs-linux-riscv64-gnu` | 0.9.1 | MIT |
| `@ironcalc/nodejs-linux-x64-gnu` | 0.9.1 | MIT |
| `@ironcalc/nodejs-win32-arm64-msvc` | 0.9.1 | MIT |
| `@ironcalc/nodejs-win32-x64-msvc` | 0.9.1 | MIT |
| `@js-sdsl/ordered-map` | 4.4.2 | MIT |
| `@protobufjs/aspromise` | 1.1.2 | BSD-3-Clause |
| `@protobufjs/base64` | 1.1.2 | BSD-3-Clause |
| `@protobufjs/codegen` | 2.0.5 | BSD-3-Clause |
| `@protobufjs/eventemitter` | 1.1.1 | BSD-3-Clause |
| `@protobufjs/fetch` | 1.1.1 | BSD-3-Clause |
| `@protobufjs/float` | 1.0.2 | BSD-3-Clause |
| `@protobufjs/path` | 1.1.2 | BSD-3-Clause |
| `@protobufjs/pool` | 1.1.0 | BSD-3-Clause |
| `@protobufjs/utf8` | 1.1.2 | BSD-3-Clause |
| `@types/node` | 26.6.6 | MIT |
| `@univerjs/core` | 1.6.0 | Apache-2.0 |
| `@univerjs/docs` | 1.6.0 | Apache-2.0 |
| `@univerjs/engine-formula` | 1.6.0 | Apache-2.0 |
| `@univerjs/engine-render` | 1.6.0 | Apache-2.0 |
| `@univerjs/icons-svg` | 1.50.0 | MIT |
| `@univerjs/protocol` | 1.6.0 | Apache-2.0 |
| `@univerjs/rpc` | 1.6.0 | Apache-2.0 |
| `@univerjs/sheets` | 1.6.0 | Apache-2.0 |
| `@univerjs/sheets-formula` | 1.6.0 | Apache-2.0 |
| `@univerjs/sheets-numfmt` | 1.6.0 | Apache-2.0 |
| `@univerjs/themes` | 1.6.0 | Apache-2.0 |
| `@wendellhu/redi` | 1.1.3 | MIT |
| `ansi-regex` | 5.0.1 | MIT |
| `ansi-styles` | 4.3.0 | MIT |
| `async-lock` | 1.5.0 | MIT |
| `cliui` | 8.0.1 | ISC |
| `collapse-white-space` | 2.1.0 | MIT |
| `color-convert` | 2.0.1 | MIT |
| `color-name` | 1.1.4 | MIT |
| `decimal.js` | 10.6.0 | MIT |
| `emoji-regex` | 8.0.0 | MIT |
| `escalade` | 3.2.0 | MIT |
| `fast-diff` | 1.3.0 | Apache-2.0 |
| `franc-min` | 6.2.0 | MIT |
| `get-caller-file` | 2.0.5 | ISC |
| `is-fullwidth-code-point` | 3.0.0 | MIT |
| `kdbush` | 4.1.0 | ISC |
| `lodash-es` | 4.18.1 | MIT |
| `lodash.camelcase` | 4.3.0 | MIT |
| `long` | 5.3.2 | Apache-2.0 |
| `n-gram` | 2.0.2 | MIT |
| `ot-json1` | 1.0.2 | ISC |
| `ot-text-unicode` | 4.0.0 | ISC |
| `protobufjs` | 7.6.6 | BSD-3-Clause |
| `quickselect` | 3.0.0 | ISC |
| `rbush` | 4.0.1 | MIT |
| `require-directory` | 2.1.1 | MIT |
| `rxjs` | 7.8.2 | Apache-2.0 |
| `string-width` | 4.2.3 | MIT |
| `strip-ansi` | 6.0.1 | MIT |
| `trigram-utils` | 2.0.1 | MIT |
| `tslib` | 2.9.1 | 0BSD |
| `undici-types` | 8.9.0 | MIT |
| `unicount` | 1.1.0 | ISC |
| `wrap-ansi` | 7.0.0 | MIT |
| `xlsx` | 0.20.3 | Apache-2.0 |
| `y18n` | 5.0.8 | ISC |
| `yargs` | 17.7.3 | MIT |
| `yargs-parser` | 21.1.1 | ISC |

## Paid packages

No paid package is installed.

- SheetJS Pro is not a dependency. The only SheetJS package is `xlsx` 0.20.3 from the Community Edition tarball above. Its README mentions SheetJS Pro as a separate product. That product is not installed.
- `@univerjs-pro` and Univer Server are not dependencies. The Univer packages are the open-source 1.6.0 packages in the table. `@univerjs/protocol` 1.6.0 (Apache-2.0) contains generated TypeScript under `lib/types/ts/univerpro/`. That directory is part of the open-source protocol package, not a separate paid package.
- HyperFormula is not installed. Its differences page is cited only as a reference for the F11–F13 formulas. See Fixtures.

## CI actions

`.github/workflows/test.yml` uses `actions/checkout@v4`, `actions/setup-python@v5`, and `actions/setup-node@v4`. Those are GitHub Actions, MIT licensed, pinned to a major tag, and not vendored.

## Fixtures

Every `.xlsx` file was written by this repository. `fixtures/build_fixtures.py` writes F01–F25 with XlsxWriter. `fixtures/build_features.py` writes `RT-features.xlsx` with XlsxWriter. `make test` does not regenerate them. Workbook creator and last-modified-by fields are empty. `Application` is `Microsoft Excel` because that is XlsxWriter's default, not because the file was saved by Excel. No upstream workbook is committed.

Expected values come from recorded desktop Excel runs. The citation for each cell is in `fixtures/fixtures.json`. Adam Marsh's desktop Excel (16.0300) run on 2026-11-23 is the primary source for every formula cell except F11–F13, which are his desktop Excel for Microsoft 365 run on 2026-11-22. The URLs below are the previous citations kept on those cells.

| File | Source of the expected values |
|---|---|
| `F01-let-lambda-call.xlsx` | https://github.com/dream-num/univer/issues/8332 |
| `F02-iso-ceiling.xlsx` | https://github.com/dream-num/univer/issues/8268; https://support.microsoft.com/en-us/excel/functions/iso-ceiling-function |
| `F03-edate-eomonth-negative-fraction.xlsx` | https://github.com/dream-num/univer/issues/8269; https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5; https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628 |
| `F04-text-negative-zero.xlsx` | https://github.com/dream-num/univer/issues/8270; https://www.mrexcel.com/board/threads/negative-number-format.120253/; https://learn.microsoft.com/en-us/answers/questions/4790299/sum-formula-returns-(0-00)-not-for-zero-sum |
| `F05-search-wildcards-and-round-digits.xlsx` | https://github.com/dream-num/univer/issues/8198; https://support.microsoft.com/en-us/office/search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495 |
| `F06-char-range.xlsx` | https://github.com/dream-num/univer/issues/8211; https://support.microsoft.com/en-us/office/char-function-bbd249c8-b36e-4a91-8017-1c133f9b837a |
| `F07-value-accounting-negatives.xlsx` | https://github.com/dream-num/univer/issues/8231 |
| `F08-amordegrc.xlsx` | https://github.com/dream-num/univer/issues/8267; https://support.microsoft.com/en-us/office/amordegrc-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51 |
| `F09-ast-cache-collision-cells.xlsx` | https://github.com/dream-num/univer/issues/7986 |
| `F10-sumifs-legacy-whole-column.xlsx` | https://github.com/dream-num/univer/issues/8154; https://support.microsoft.com/en-us/office/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34 |
| `F11-hf-negative-rounding.xlsx` | Adam Marsh, desktop Excel for Microsoft 365, 2026-11-22. Reference: https://hyperformula.handsontable.com/docs/guide/list-of-differences; https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef; https://support.microsoft.com/en-us/office/iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356 |
| `F12-hf-npv-sequence-errors.xlsx` | Adam Marsh, desktop Excel for Microsoft 365, 2026-11-22. Reference: https://hyperformula.handsontable.com/docs/guide/list-of-differences; https://support.microsoft.com/en-us/office/t-inv-function-2908272b-4e61-4942-9df9-a25fec9b0e2e; https://learn.microsoft.com/en-us/answers/questions/5065194/weibull-dist-function-error |
| `F13-hf-if-false-cycle.xlsx` | Adam Marsh, desktop Excel for Microsoft 365, 2026-11-22. Reference: https://hyperformula.handsontable.com/docs/guide/list-of-differences |
| `F14-large-k.xlsx` | https://github.com/ironcalc/IronCalc/issues/1750 |
| `F15-vlookup-index-below-one.xlsx` | https://github.com/ironcalc/IronCalc/issues/1751 |
| `F16-text-general.xlsx` | https://github.com/ironcalc/IronCalc/issues/1752 |
| `F17-explicit-argument-coercion.xlsx` | https://github.com/ironcalc/IronCalc/issues/1605; https://support.microsoft.com/en-us/office/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0; https://support.microsoft.com/en-us/office/var-function-1f2b7ab2-954d-4e17-ba2c-9e58b15a7da2 |
| `F18-xlookup-return-and-empty.xlsx` | https://github.com/ironcalc/IronCalc/issues/1789; https://github.com/ironcalc/IronCalc/issues/1718 |
| `F19-sort-empty-argument.xlsx` | https://github.com/ironcalc/IronCalc/issues/1597 |
| `F20-date-serial-zero.xlsx` | https://github.com/ironcalc/IronCalc/issues/1569 |
| `F21-date-year-offset.xlsx` | https://github.com/ironcalc/IronCalc/issues/1364 |
| `F22-and-empty-text.xlsx` | https://github.com/ironcalc/IronCalc/issues/1609 |
| `F23-sum-if-range.xlsx` | https://github.com/ironcalc/IronCalc/issues/1749 |
| `F24-subtotal-edges.xlsx` | https://github.com/ironcalc/IronCalc/issues/1739 |
| `F25-decimal-leading-sign.xlsx` | https://github.com/ironcalc/IronCalc/issues/1615 |
| `RT-features.xlsx` | Written by `fixtures/build_features.py`. No external workbook. |

### Source notes

Univer issues (`dream-num/univer`) and IronCalc issues (`ironcalc/IronCalc`) are cited for the formula and the Excel result the issue states. The issue text is not copied into a workbook. IronCalc itself is MIT OR Apache-2.0. A GitHub issue is not automatically under that license. What this repo keeps is the short formula and the stated result, regenerated into an xlsx by the build script.

F17's note says the Excel values in IronCalc issue 1605 were cached results from an Excel-saved Apache POI workbook (`EverythingTests`). Apache POI is Apache-2.0. That workbook is not in this repository. The fixture was regenerated here from the formulas and results in the issue.

Microsoft Support and Microsoft Learn pages are cited by URL for stated function results. F08 uses the inputs and result 776 from the published AMORDEGRC example. The documentation text is not copied.

F04's expected value is `"0"` from Adam Marsh, desktop Excel (16.0300), 2026-11-23. Univer issue 8270, a MrExcel thread, and a Microsoft Q&A page are the previous citations. The issue author said they could not run the formula in Excel. Neither post nor any attachment is in this repository.

F11, F12, and F13 expected values come from our own Excel run. Adam Marsh typed each formula into desktop Excel for Microsoft 365 on 2026-11-22, and each result matched the value already stored for that cell. That measurement is original. The other formula cells were remeasured by Adam Marsh in desktop Excel (16.0300) on 2026-11-23. F18 `D1` was entered as a dynamic-array formula, with E1 cleared, and the first spill cell was `x`. [HyperFormula's list of differences](https://hyperformula.handsontable.com/docs/guide/list-of-differences) is cited only as a reference for F11–F13. Nothing is copied from that page beyond the short formulas. HyperFormula is not installed and not vendored.

