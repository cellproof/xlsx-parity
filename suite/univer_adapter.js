// Headless Univer 1.6.0. Two passes:
//
// 1. Engine hook. UniverFormulaEnginePlugin._initializeWithOverride is the
//    protected binding point (published in plugin.d.ts). It replaces
//    ICalculateFormulaService with a CalculateFormulaService subclass. The
//    subclass awaits FormulaResultService.applyResult, which is the contract
//    in the engine-formula README. Cell values are computed by @ironcalc/nodejs.
//
// 2. Workbook ingest. SheetJS CE reads the xlsx into Univer IWorkbookData.
//    ICellData.f is the formula text ("=LET(...)"). OOXML storage prefixes
//    (_xlfn. / _xlpm.) are stripped into that field. The native pass uses
//    Univer's own formula engine and does not install the IronCalc subclass.
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const {
  CellValueType,
  IConfigService,
  IUniverInstanceService,
  LocaleType,
  ObjectMatrix,
  Univer,
  UniverInstanceType,
  mergeLocales,
} = require("@univerjs/core");
const { FUniver } = require("@univerjs/core/facade");
const { UniverDocsPlugin } = require("@univerjs/docs");
const {
  AstTreeBuilder,
  CalculateFormulaService,
  FormulaExecutedStateType,
  FormulaResultService,
  ICalculateFormulaService,
  IFormulaCurrentConfigService,
  IFormulaDependencyGenerator,
  IFormulaRuntimeService,
  Interpreter,
  Lexer,
  UniverFormulaEnginePlugin,
} = require("@univerjs/engine-formula");
const { UniverRenderEnginePlugin } = require("@univerjs/engine-render");
const { UniverSheetsPlugin } = require("@univerjs/sheets");
const { UniverSheetsFormulaPlugin } = require("@univerjs/sheets-formula");
const { UniverSheetsNumfmtPlugin } = require("@univerjs/sheets-numfmt");
const SheetsEnUS = require("@univerjs/sheets/locale/en-US");
const { Model } = require("@ironcalc/nodejs");

require("@univerjs/engine-formula/facade");
require("@univerjs/sheets/facade");
require("@univerjs/sheets-formula/facade");
require("@univerjs/sheets-numfmt/facade");

const HOOK_FIXTURE = "F01-let-lambda-call";
const PACKAGES = [
  "@ironcalc/nodejs",
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

const root = path.resolve(__dirname, "..");

function packageVersion(name) {
  const pkgPath = path.join(root, "node_modules", ...name.split("/"), "package.json");
  return JSON.parse(fs.readFileSync(pkgPath, "utf8")).version;
}

function formulaText(stored) {
  const body = String(stored).replace(/_xl(fn|pm|ws)\./g, "");
  return body.startsWith("=") ? body : `=${body}`;
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

function cellType(value) {
  if (typeof value === "number") return CellValueType.NUMBER;
  if (typeof value === "boolean") return CellValueType.BOOLEAN;
  return CellValueType.STRING;
}

function walkMatrix(matrix, fn) {
  if (!matrix) return;
  if (typeof matrix.forValue === "function") {
    matrix.forValue((row, col, value) => fn(row, col, value));
    return;
  }
  for (const rowKey of Object.keys(matrix)) {
    const row = matrix[rowKey];
    if (!row || typeof row !== "object") continue;
    for (const colKey of Object.keys(row)) fn(Number(rowKey), Number(colKey), row[colKey]);
  }
}

function sheetJsToWorkbookData(wb, id) {
  const sheetOrder = [];
  const sheets = {};
  wb.SheetNames.forEach((name, index) => {
    const sheetId = `sheet-${index + 1}`;
    sheetOrder.push(sheetId);
    const src = wb.Sheets[name] || {};
    const cellData = {};
    let rowCount = 1;
    let columnCount = 1;
    if (src["!ref"]) {
      const range = XLSX.utils.decode_range(src["!ref"]);
      rowCount = range.e.r + 1;
      columnCount = range.e.c + 1;
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = src[XLSX.utils.encode_cell({ r, c })];
          if (!cell || cell.t === "z") continue;
          const out = {};
          if (cell.f) out.f = formulaText(cell.f);
          else if (cell.t === "n" || cell.t === "b" || cell.t === "s" || cell.t === "str") out.v = cell.v;
          else if (cell.t === "e") out.v = cell.w;
          else continue;
          if (!cellData[r]) cellData[r] = {};
          cellData[r][c] = out;
        }
      }
    }
    sheets[sheetId] = { id: sheetId, name, cellData, rowCount, columnCount };
  });
  return {
    id,
    name: id,
    appVersion: "1.6.0",
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder,
    sheets,
  };
}

function readFixtureWorkbook(fx) {
  const abs = path.join(root, "fixtures", fx.file);
  const xlsx = XLSX.readFile(abs, { cellFormula: true, xlfn: true, cellDates: false });
  return sheetJsToWorkbookData(xlsx, fx.id);
}

class IronCalcCalculateFormulaService extends CalculateFormulaService {
  constructor(config, lexer, current, runtime, dep, interpreter, ast, result, univer) {
    super(config, lexer, current, runtime, dep, interpreter, ast, result);
    this._univerInstances = univer;
  }

  async execute(config) {
    if (this._disposed) return;
    const model = new Model("hook", "en", "UTC", "en");
    const sheetIndex = new Map();
    let nextIndex = 0;
    for (const unitId of Object.keys(config.formulaData || {})) {
      const workbook = this._univerInstances.getUnit(unitId, UniverInstanceType.UNIVER_SHEET);
      const sheets = workbook ? workbook.getSheets() : [];
      for (const sheet of sheets) {
        if (nextIndex === 0) model.renameSheet(0, sheet.getName());
        else model.addSheet(sheet.getName());
        const index = nextIndex;
        nextIndex += 1;
        sheetIndex.set(`${unitId}\0${sheet.getSheetId()}`, index);
        sheet.getCellMatrix().forValue((row, col, cell) => {
          if (!cell) return;
          if (cell.f) model.updateCellWithFormula(index, row + 1, col + 1, formulaText(cell.f));
          else if (typeof cell.v === "number") model.updateCellWithNumber(index, row + 1, col + 1, cell.v);
          else if (typeof cell.v === "boolean") model.updateCellWithBool(index, row + 1, col + 1, cell.v);
          else if (cell.v != null) model.updateCellWithText(index, row + 1, col + 1, String(cell.v));
        });
      }
    }
    model.evaluate();
    const unitData = Object.create(null);
    for (const unitId of Object.keys(config.formulaData || {})) {
      unitData[unitId] = Object.create(null);
      const formulaSheets = config.formulaData[unitId] || {};
      for (const sheetId of Object.keys(formulaSheets)) {
        const index = sheetIndex.get(`${unitId}\0${sheetId}`);
        const matrix = new ObjectMatrix();
        walkMatrix(formulaSheets[sheetId], (row, col) => {
          if (index == null) return;
          const value = model.getCellValue(index, row + 1, col + 1);
          matrix.setValue(row, col, { v: value, t: cellType(value) });
        });
        unitData[unitId][sheetId] = matrix;
      }
    }
    const data = {
      functionsExecutedState: FormulaExecutedStateType.SUCCESS,
      unitData,
      arrayFormulaRange: Object.create(null),
      arrayFormulaEmbedded: Object.create(null),
      unitOtherData: Object.create(null),
      arrayFormulaCellData: Object.create(null),
      clearArrayFormulaCellData: Object.create(null),
      imageFormulaData: [],
      runtimeFeatureRange: Object.create(null),
      runtimeFeatureCellData: Object.create(null),
      dependencyTreeModelData: [],
    };
    await this._formulaResultService.applyResult(data);
    if (!this._disposed) this._executionCompleteListener$.next(data);
  }
}

class IronCalcFormulaPlugin extends UniverFormulaEnginePlugin {
  _initializeWithOverride() {
    super._initializeWithOverride();
    this._injector.replace([
      ICalculateFormulaService,
      {
        useFactory: (config, lexer, current, runtime, dep, interpreter, ast, result, univer) =>
          new IronCalcCalculateFormulaService(
            config, lexer, current, runtime, dep, interpreter, ast, result, univer,
          ),
        deps: [
          IConfigService,
          Lexer,
          IFormulaCurrentConfigService,
          IFormulaRuntimeService,
          IFormulaDependencyGenerator,
          Interpreter,
          AstTreeBuilder,
          FormulaResultService,
          IUniverInstanceService,
        ],
      },
    ]);
  }
}

function registerHeadless(univer, formulaPlugin) {
  univer.registerPlugin(UniverRenderEnginePlugin);
  univer.registerPlugin(formulaPlugin);
  univer.registerPlugin(UniverDocsPlugin);
  univer.registerPlugin(UniverSheetsPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin);
  univer.registerPlugin(UniverSheetsNumfmtPlugin);
}

async function calculate(snapshot, formulaPlugin) {
  const univer = new Univer({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(SheetsEnUS) },
  });
  try {
    registerHeadless(univer, formulaPlugin);
    const api = FUniver.newAPI(univer);
    const workbook = api.createWorkbook(snapshot);
    await api.getFormula().onCalculationResultApplied(30000);
    return { univer, workbook };
  } catch (exc) {
    univer.dispose();
    throw exc;
  }
}

function readCases(fx, workbook, modelError) {
  const cases = [];
  for (const [sheetName, addr] of fx.cases) {
    const ref = `${sheetName}!${addr}`;
    if (modelError) {
      cases.push({ file: fx.file, cell: ref, got: null, error: modelError });
      continue;
    }
    try {
      const sheet = workbook.getSheetByName(sheetName);
      if (!sheet) throw new Error(`missing sheet ${sheetName}`);
      cases.push({ file: fx.file, cell: ref, got: jsonValue(sheet.getRange(addr).getValue()), error: null });
    } catch (exc) {
      cases.push({ file: fx.file, cell: ref, got: null, error: errorText(exc) });
    }
  }
  return cases;
}

async function runFixture(fx, formulaPlugin) {
  const snapshot = readFixtureWorkbook(fx);
  let univer = null;
  let workbook = null;
  let modelError = null;
  try {
    const ran = await calculate(snapshot, formulaPlugin);
    univer = ran.univer;
    workbook = ran.workbook;
  } catch (exc) {
    modelError = errorText(exc);
  }
  const cases = readCases(fx, workbook, modelError);
  if (univer) univer.dispose();
  return cases;
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const fixtures = JSON.parse(fs.readFileSync(path.join(root, "fixtures", "fixtures.json"), "utf8"));
  const packages = {};
  for (const name of PACKAGES) packages[name] = packageVersion(name);

  const native = [];
  for (const fx of fixtures) native.push(...await runFixture(fx, UniverFormulaEnginePlugin));

  const hookFx = fixtures.find((item) => item.id === HOOK_FIXTURE);
  if (!hookFx) {
    console.error(`missing fixture ${HOOK_FIXTURE}`);
    process.exit(1);
  }
  const hook = await runFixture(hookFx, IronCalcFormulaPlugin);

  const report = { packages, hook, native };
  fs.writeFileSync(path.join(outDir, "read.json"), JSON.stringify(report));
}

main().catch((exc) => {
  console.error(errorText(exc));
  process.exit(1);
});
