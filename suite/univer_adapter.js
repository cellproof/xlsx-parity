// Headless Univer calculates fixture F01. Plugin registration follows the
// open-source Node guide: core, engine-render, engine-formula, docs, sheets,
// sheets-formula, sheets-numfmt. No *-ui package and no @univerjs-pro package.
//
// Excel stores the formulas with OOXML prefixes (_xlfn. / _xlpm.). Univer's
// facade calculates formula text (`=LET(...)`). The adapter strips those
// prefixes the same way suite/run.py human() displays a formula.
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { LocaleType, mergeLocales, Univer } = require("@univerjs/core");
const { FUniver } = require("@univerjs/core/facade");
const { UniverDocsPlugin } = require("@univerjs/docs");
const { UniverFormulaEnginePlugin } = require("@univerjs/engine-formula");
const { UniverRenderEnginePlugin } = require("@univerjs/engine-render");
const { UniverSheetsPlugin } = require("@univerjs/sheets");
const { UniverSheetsFormulaPlugin } = require("@univerjs/sheets-formula");
const { UniverSheetsNumfmtPlugin } = require("@univerjs/sheets-numfmt");
const SheetsEnUS = require("@univerjs/sheets/locale/en-US");

require("@univerjs/engine-formula/facade");
require("@univerjs/sheets/facade");
require("@univerjs/sheets-formula/facade");
require("@univerjs/sheets-numfmt/facade");

const FIXTURE_ID = "F01-let-lambda-call";
const PACKAGES = [
  "@univerjs/core",
  "@univerjs/docs",
  "@univerjs/engine-formula",
  "@univerjs/engine-render",
  "@univerjs/sheets",
  "@univerjs/sheets-formula",
  "@univerjs/sheets-numfmt",
  "rxjs",
];

const outDir = process.argv[2];
if (!outDir) {
  console.error("usage: univer_adapter.js <out-dir>");
  process.exit(2);
}

function packageVersion(name) {
  const pkgPath = path.join(root, "node_modules", ...name.split("/"), "package.json");
  return JSON.parse(fs.readFileSync(pkgPath, "utf8")).version;
}

function formulaText(stored) {
  const body = String(stored).replace(/_xl(fn|pm|ws)\./g, "");
  return body.startsWith("=") ? body : `=${body}`;
}

function cellIndex(addr) {
  const match = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!match) {
    throw new Error(`bad cell ${addr}`);
  }
  let col = 0;
  for (const ch of match[1]) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { row: Number(match[2]) - 1, col: col - 1 };
}

function jsonValue(value) {
  if (value === undefined) return null;
  if (typeof value === "number" || typeof value === "string" || typeof value === "boolean" || value === null) {
    return value;
  }
  return String(value);
}

function errorText(exc) {
  return exc instanceof Error ? `${exc.name}: ${exc.message}` : String(exc);
}

const root = path.resolve(__dirname, "..");

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const fixtures = JSON.parse(fs.readFileSync(path.join(root, "fixtures", "fixtures.json"), "utf8"));
  const fx = fixtures.find((item) => item.id === FIXTURE_ID);
  if (!fx) {
    console.error(`missing fixture ${FIXTURE_ID}`);
    process.exit(1);
  }

  const abs = path.join(root, "fixtures", fx.file);
  const xlsx = XLSX.readFile(abs, { cellFormula: true, xlfn: true });
  const loaded = [];
  for (const [sheet, addr, stored] of fx.cases) {
    const cell = xlsx.Sheets[sheet] && xlsx.Sheets[sheet][addr];
    if (!cell || cell.f !== stored) {
      loaded.push({
        file: fx.file,
        cell: `${sheet}!${addr}`,
        got: null,
        error: `xlsx formula ${cell && cell.f} != fixtures.json ${stored}`,
      });
      continue;
    }
    loaded.push({ file: fx.file, sheet, addr, cell: `${sheet}!${addr}`, formula: formulaText(stored), error: null });
  }

  const packages = {};
  for (const name of PACKAGES) packages[name] = packageVersion(name);

  const runnable = loaded.filter((item) => !item.error);
  if (runnable.length) {
    let modelError = null;
    let workbook = null;
    const univer = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: mergeLocales(SheetsEnUS) },
    });
    try {
      univer.registerPlugin(UniverRenderEnginePlugin);
      univer.registerPlugin(UniverFormulaEnginePlugin);
      univer.registerPlugin(UniverDocsPlugin);
      univer.registerPlugin(UniverSheetsPlugin);
      univer.registerPlugin(UniverSheetsFormulaPlugin);
      univer.registerPlugin(UniverSheetsNumfmtPlugin);
      const api = FUniver.newAPI(univer);
      const sheets = {};
      const order = [];
      for (const item of runnable) {
        if (!sheets[item.sheet]) {
          const id = `sheet-${order.length + 1}`;
          order.push(id);
          sheets[item.sheet] = { id, name: item.sheet, cellData: {} };
        }
        const { row, col } = cellIndex(item.addr);
        const data = sheets[item.sheet].cellData;
        if (!data[row]) data[row] = {};
        data[row][col] = { f: item.formula };
      }
      const sheetOrder = order;
      const sheetMap = {};
      for (const name of Object.keys(sheets)) sheetMap[sheets[name].id] = sheets[name];
      workbook = api.createWorkbook({ id: "f01", name: "F01", sheetOrder, sheets: sheetMap });
      await api.getFormula().onCalculationResultApplied(15000);
    } catch (exc) {
      modelError = errorText(exc);
    }
    for (const item of runnable) {
      if (modelError || !workbook) {
        item.got = null;
        item.error = modelError || "workbook was not created";
        continue;
      }
      try {
        const sheet = workbook.getSheetByName(item.sheet);
        if (!sheet) throw new Error(`missing sheet ${item.sheet}`);
        item.got = jsonValue(sheet.getRange(item.addr).getValue());
      } catch (exc) {
        item.got = null;
        item.error = errorText(exc);
      }
    }
    univer.dispose();
  }

  const cases = loaded.map((item) => ({
    file: item.file,
    cell: item.cell,
    got: item.got === undefined ? null : item.got,
    error: item.error,
  }));
  fs.writeFileSync(path.join(outDir, "read.json"), JSON.stringify({ packages, cases }));
}

main().catch((exc) => {
  console.error(errorText(exc));
  process.exit(1);
});
