import XLSX, { version as SHEETJS_VERSION } from "xlsx";
import { excelMatch } from "./compare.js";

// Scientific format with 16 digits after the decimal. IronCalc's General
// format rounds (1/33600 displays as 0.000029762), which is coarser than
// the suite's 1e-12 relative tolerance. This format round-trips the double
// well inside that tolerance. @ironcalc/wasm has no raw-value getter;
// getFormattedCellValue is the typed result after this format is applied.
const PRECISE_FMT = "0.0000000000000000E+00";

const CELL_NUMBER = 1;
const CELL_TEXT = 2;
const CELL_LOGICAL = 4;
const CELL_ERROR = 16;

export const LOADER_SUMMARY =
  "Formulas loaded via SheetJS into @ironcalc/wasm Model.setUserInput. This package does not export Model.fromXlsx.";

export function loaderDetail(sheetjsVersion, ironVersion) {
  return (
    `@ironcalc/wasm ${ironVersion} does not export Model.fromXlsx. ` +
    `The installed declaration file exposes Model.fromBytes, setUserInput, and getFormattedCellValue. ` +
    `Model.fromXlsx is compiled only into @ironcalc/wasm-xlsx. ` +
    `This page reads the file with SheetJS CE ${sheetjsVersion} and enters every cell with setUserInput. ` +
    `OOXML storage prefixes (_xlfn., _xlpm., _xlws.) are removed first, because setUserInput expects the formula a person would type. ` +
    `Number formats from the file are not copied, so a date serial stays a number. ` +
    `The xlsx importer's implicit-intersection rewrite is not applied. ` +
    `That is a real difference from the scoreboard: @ironcalc/nodejs Model.fromXlsx records F23 SUM(IF(...)) as #VALUE!, ` +
    `while setUserInput on those same formulas returns Excel's cached 17 and 100. ` +
    `The scoreboard is the fromXlsx baseline. This scanner is the setUserInput path.`
  );
}

function userFormula(stored) {
  const stripped = stored.replace(/_xl(fn|pm|ws)\./g, "");
  const input = stripped.startsWith("=") ? stripped : `=${stripped}`;
  return { input, strippedPrefixes: stripped !== stored };
}

function textInput(value) {
  // A leading apostrophe forces text and is not part of the value,
  // including for "", "TRUE", "123", and strings that start with "=".
  return `'${value}`;
}

function numberInput(value) {
  if (Object.is(value, -0)) return "0";
  return String(value);
}

function cachedExcel(cell) {
  if (!cell) return { present: false, value: null };
  if (cell.t === "e") {
    if (typeof cell.w !== "string" || cell.w === "") return { present: false, value: null };
    return { present: true, value: cell.w };
  }
  if (cell.v === undefined) return { present: false, value: null };
  return { present: true, value: cell.v };
}

function hiddenSpans(entries) {
  const spans = [];
  let start = null;
  const length = entries ? entries.length : 0;
  for (let idx = 0; idx < length; idx++) {
    const hidden = !!(entries[idx] && entries[idx].hidden);
    if (hidden) {
      if (start == null) start = idx + 1;
    } else if (start != null) {
      spans.push([start, idx]);
      start = null;
    }
  }
  if (start != null) spans.push([start, length]);
  return spans;
}

function readCalculated(model, sheet, row, column) {
  const type = model.getCellType(sheet, row, column);
  const formatted = model.getFormattedCellValue(sheet, row, column);
  if (type === CELL_ERROR) return formatted;
  if (type === CELL_LOGICAL) {
    if (formatted === "TRUE") return true;
    if (formatted === "FALSE") return false;
    throw new Error(`logical cell formatted as ${JSON.stringify(formatted)}`);
  }
  if (type === CELL_TEXT) return formatted;
  if (type === CELL_NUMBER) {
    if (formatted === "") return null;
    model.updateRangeStyle(
      { sheet, row, column, width: 1, height: 1 },
      "num_fmt",
      PRECISE_FMT,
    );
    const precise = model.getFormattedCellValue(sheet, row, column).trim();
    const parsed = Number(precise);
    if (!Number.isFinite(parsed)) {
      throw new Error(`number format ${JSON.stringify(precise)} did not parse`);
    }
    return parsed;
  }
  // Array (64) and compound (128): the wasm binding exposes the anchor's
  // formatted text, not a raw array. Compare that text.
  return formatted;
}

function cellInput(cell) {
  if (typeof cell.f === "string" && cell.f !== "") return userFormula(cell.f);
  if (cell.t === "b") return { input: cell.v ? "TRUE" : "FALSE", strippedPrefixes: false };
  if (cell.t === "e") return { input: cell.w || "#N/A", strippedPrefixes: false };
  if (cell.t === "s" || cell.t === "str") return { input: textInput(cell.v == null ? "" : String(cell.v)), strippedPrefixes: false };
  if (cell.t === "n" && typeof cell.v === "number") return { input: numberInput(cell.v), strippedPrefixes: false };
  if (cell.v == null) return null;
  return { input: textInput(String(cell.v)), strippedPrefixes: false };
}

export function scanWorkbook(model, bytes, fileName) {
  if (SHEETJS_VERSION !== "0.20.3") {
    throw new Error(`SheetJS version is ${SHEETJS_VERSION}, expected 0.20.3`);
  }
  const workbook = XLSX.read(bytes, {
    type: "array",
    cellFormula: true,
    xlfn: true,
    cellDates: false,
    cellNF: false,
    cellStyles: false,
  });
  const names = workbook.SheetNames || [];
  if (names.length === 0) throw new Error("Workbook has no sheets");

  const warnings = [];
  let strippedPrefixCount = 0;
  model.pauseEvaluation();
  model.renameSheet(0, "~parity~0");
  for (let i = 1; i < names.length; i++) {
    model.newSheet();
    model.renameSheet(i, `~parity~${i}`);
  }
  for (let i = 0; i < names.length; i++) {
    model.renameSheet(i, names[i]);
  }

  const defined = workbook.Workbook && workbook.Workbook.Names;
  if (Array.isArray(defined)) {
    for (const entry of defined) {
      if (!entry || typeof entry.Name !== "string" || typeof entry.Ref !== "string" || entry.Ref === "") continue;
      const formula = entry.Ref.startsWith("=") ? entry.Ref : `=${entry.Ref}`;
      const scope = typeof entry.Sheet === "number" ? entry.Sheet : null;
      try {
        model.newDefinedName(entry.Name, scope, formula);
      } catch (err) {
        warnings.push(`Defined name ${entry.Name} was not loaded: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  }

  const formulaCells = [];
  for (let s = 0; s < names.length; s++) {
    const sheet = workbook.Sheets[names[s]];
    for (const span of hiddenSpans(sheet["!rows"])) {
      model.setRowsHidden(s, span[0], span[1], true);
    }
    for (const span of hiddenSpans(sheet["!cols"])) {
      model.setColumnsHidden(s, span[0], span[1], true);
    }
    for (const addr of Object.keys(sheet)) {
      if (addr.startsWith("!")) continue;
      const cell = sheet[addr];
      if (!cell) continue;
      const decoded = XLSX.utils.decode_cell(addr);
      const row = decoded.r + 1;
      const column = decoded.c + 1;
      const spec = cellInput(cell);
      if (!spec) continue;
      if (spec.strippedPrefixes) strippedPrefixCount += 1;
      try {
        model.setUserInput(s, row, column, spec.input);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        if (typeof cell.f === "string" && cell.f !== "") {
          formulaCells.push({
            sheet: names[s],
            cell: addr,
            formula: cell.f,
            excel: cachedExcel(cell),
            iron: { ok: false, value: null, error: message },
          });
        } else {
          warnings.push(`${names[s]}!${addr} was not entered: ${message}`);
        }
        continue;
      }
      if (typeof cell.f === "string" && cell.f !== "") {
        formulaCells.push({
          sheet: names[s],
          cell: addr,
          formula: cell.f,
          excel: cachedExcel(cell),
          iron: null,
          at: [s, row, column],
        });
      }
    }
  }

  model.resumeEvaluation();
  model.evaluate();

  const rows = [];
  for (const item of formulaCells) {
    let ironValue = null;
    let ironError = null;
    if (item.iron && item.iron.ok === false) {
      ironError = item.iron.error;
    } else {
      try {
        ironValue = readCalculated(model, item.at[0], item.at[1], item.at[2]);
      } catch (err) {
        ironError = err instanceof Error ? err.message : String(err);
      }
    }
    let status;
    if (!item.excel.present) status = "uncached";
    else if (ironError) status = "error";
    else if (excelMatch(ironValue, item.excel.value)) status = "match";
    else status = "mismatch";
    rows.push({
      sheet: item.sheet,
      cell: item.cell,
      formula: item.formula,
      excel: item.excel.present ? item.excel.value : null,
      excelCached: item.excel.present,
      ironcalc: ironError ? null : ironValue,
      error: ironError,
      status,
    });
  }

  if (strippedPrefixCount > 0) {
    warnings.push(
      `${strippedPrefixCount} formula${strippedPrefixCount === 1 ? "" : "s"} had _xlfn. / _xlpm. / _xlws. removed before setUserInput. The table still shows the stored formula.`,
    );
  }

  rows.sort((a, b) => {
    const sheetOrder = names.indexOf(a.sheet) - names.indexOf(b.sheet);
    if (sheetOrder !== 0) return sheetOrder;
    const left = XLSX.utils.decode_cell(a.cell);
    const right = XLSX.utils.decode_cell(b.cell);
    if (left.r !== right.r) return left.r - right.r;
    return left.c - right.c;
  });

  const summary = {
    formulaCells: rows.length,
    match: rows.filter((row) => row.status === "match").length,
    mismatch: rows.filter((row) => row.status === "mismatch").length,
    uncached: rows.filter((row) => row.status === "uncached").length,
    error: rows.filter((row) => row.status === "error").length,
  };

  return {
    file: fileName,
    sheetjs: SHEETJS_VERSION,
    loader: "sheetjs-setUserInput",
    nativeXlsxImport: false,
    summary,
    warnings,
    cells: rows,
  };
}
