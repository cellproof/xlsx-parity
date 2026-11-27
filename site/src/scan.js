import XLSX, { version as SHEETJS_VERSION } from "xlsx";
import { Model } from "@ironcalc/wasm-xlsx";
import { excelMatch } from "./compare.js";

// Scientific format with 16 digits after the decimal. IronCalc's General
// format rounds (1/33600 displays as 0.000029762), which is coarser than
// the suite's 1e-12 relative tolerance. @ironcalc/wasm-xlsx exposes
// getFormattedCellValue, not a raw getter, so numeric results are read
// back through this format.
const PRECISE_FMT = "0.0000000000000000E+00";

const CELL_NUMBER = 1;
const CELL_TEXT = 2;
const CELL_LOGICAL = 4;
const CELL_ERROR = 16;

export const LOADER_SUMMARY =
  "Excel's cached values come from SheetJS CE. Formulas are calculated by @ironcalc/wasm-xlsx Model.fromXlsx and evaluate().";

export function loaderDetail(sheetjsVersion, ironVersion) {
  return (
    `SheetJS CE ${sheetjsVersion} reads each formula string and Excel's cached value. ` +
    `@ironcalc/wasm-xlsx ${ironVersion} loads the same bytes with Model.fromXlsx(bytes, name, "en", "UTC", "en") and evaluate(). ` +
    `Those are the same locale, timezone, and language arguments the suite passes to @ironcalc/nodejs Model.fromXlsx. ` +
    `Numeric results are read back through a 16-digit scientific format because this package exposes getFormattedCellValue. ` +
    `The comparison then uses the suite rule: relative tolerance 1e-12, absolute tolerance 0, and exact strings, booleans, and errors.`
  );
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
  return formatted;
}

export function scanWorkbook(bytes, fileName) {
  if (SHEETJS_VERSION !== "0.20.3") {
    throw new Error(`SheetJS version is ${SHEETJS_VERSION}, expected 0.20.3`);
  }
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const workbook = XLSX.read(data, {
    type: "array",
    cellFormula: true,
    xlfn: true,
    cellDates: false,
    cellNF: false,
    cellStyles: false,
  });
  const names = workbook.SheetNames || [];
  if (names.length === 0) throw new Error("Workbook has no sheets");

  const model = Model.fromXlsx(new Uint8Array(data), fileName, "en", "UTC", "en");
  try {
    model.evaluate();
    const props = model.getWorksheetsProperties();
    const sheetIndex = new Map();
    props.forEach((prop, index) => {
      if (prop && typeof prop.name === "string" && !sheetIndex.has(prop.name)) sheetIndex.set(prop.name, index);
    });

    const warnings = [];
    const rows = [];
    for (const name of names) {
      const sheet = workbook.Sheets[name];
      const index = sheetIndex.get(name);
      if (index == null) warnings.push(`Sheet ${name} was not in the workbook Model.fromXlsx loaded.`);
      for (const addr of Object.keys(sheet)) {
        if (addr.startsWith("!")) continue;
        const cell = sheet[addr];
        if (!cell || typeof cell.f !== "string" || cell.f === "") continue;
        const excel = cachedExcel(cell);
        const decoded = XLSX.utils.decode_cell(addr);
        let ironValue = null;
        let ironError = null;
        if (index == null) {
          ironError = `Sheet ${name} was not loaded`;
        } else {
          try {
            ironValue = readCalculated(model, index, decoded.r + 1, decoded.c + 1);
          } catch (err) {
            ironError = err instanceof Error ? err.message : String(err);
          }
        }
        let status;
        if (!excel.present) status = "uncached";
        else if (ironError) status = "error";
        else if (excelMatch(ironValue, excel.value)) status = "match";
        else status = "mismatch";
        rows.push({
          sheet: name,
          cell: addr,
          formula: cell.f,
          excel: excel.present ? excel.value : null,
          excelCached: excel.present,
          ironcalc: ironError ? null : ironValue,
          error: ironError,
          status,
        });
      }
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
      loader: "fromXlsx",
      nativeXlsxImport: true,
      summary,
      warnings,
      cells: rows,
    };
  } finally {
    model.free();
  }
}
