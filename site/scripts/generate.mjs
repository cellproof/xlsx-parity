// Build the scoreboard JSON from suite/baseline.json and fixtures/fixtures.json.
// Counts are computed with the same excel_match rule as suite/run.py, then
// checked against the summary fields in the baseline. Nothing is typed in.
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { excelMatch, showValue, valuesEqual } from "../src/compare.js";

const here = dirname(fileURLToPath(import.meta.url));
const siteRoot = join(here, "..");
const repoRoot = join(siteRoot, "..");
const baseline = JSON.parse(readFileSync(join(repoRoot, "suite", "baseline.json"), "utf8"));
const fixtures = JSON.parse(readFileSync(join(repoRoot, "fixtures", "fixtures.json"), "utf8"));

const SAMPLES = [
  "F14-large-k.xlsx",
  "F23-sum-if-range.xlsx",
  "F10-sumifs-legacy-whole-column.xlsx",
  "F12-hf-npv-sequence-errors.xlsx",
];

function die(message) {
  console.error(message);
  process.exit(1);
}

function percent(passed, total) {
  if (passed === total) return "100%";
  const tenths = Math.floor((passed * 1000 + Math.floor(total / 2)) / total);
  return `${Math.floor(tenths / 10)}.${tenths % 10}%`;
}

function loadCells() {
  const cells = [];
  const seen = new Set();
  for (const fixture of fixtures) {
    for (const entry of fixture.cases) {
      if (entry.length !== 5) die(`${fixture.id} case does not have 5 fields`);
      const [sheet, addr, formula, expected, source] = entry;
      const cell = `${sheet}!${addr}`;
      const key = `${fixture.id} ${cell}`;
      if (seen.has(key)) die(`duplicate ${key}`);
      seen.add(key);
      cells.push({ fixture: fixture.id, cell, formula, expected, source });
    }
  }
  if (cells.length === 0) die("no fixture cells");
  if (cells.length !== baseline.cells) {
    die(`fixture cells ${cells.length} != baseline.cells ${baseline.cells}`);
  }
  return cells;
}

function indexResults(results, label) {
  const found = new Map();
  if (!results || typeof results !== "object") die(`${label} results missing`);
  for (const [fixture, cells] of Object.entries(results)) {
    for (const [cell, actual] of Object.entries(cells)) {
      const key = `${fixture} ${cell}`;
      if (found.has(key)) die(`${label} repeats ${key}`);
      found.set(key, actual);
    }
  }
  return found;
}

function indexFailures(failures, label) {
  const found = new Map();
  for (const [fixture, cells] of Object.entries(failures || {})) {
    for (const [cell, actual] of Object.entries(cells)) {
      const key = `${fixture} ${cell}`;
      if (found.has(key)) die(`${label} repeats ${key}`);
      found.set(key, actual);
    }
  }
  return found;
}

function cover(found, cells, label) {
  const wanted = new Set(cells.map((cell) => `${cell.fixture} ${cell.cell}`));
  for (const key of wanted) {
    if (!found.has(key)) die(`${label} missing ${key}`);
  }
  for (const key of found.keys()) {
    if (!wanted.has(key)) die(`${label} has extra ${key}`);
  }
}

function univerVersion(packages) {
  const versions = [...new Set(Object.entries(packages).filter(([name]) => name.startsWith("@univerjs/")).map(([, version]) => version))];
  if (versions.length !== 1) die(`Univer packages do not share one version: ${versions.join(", ")}`);
  return versions[0];
}

function libreofficeVersion(spec) {
  const match = /^LibreOffice (\d+(?:\.\d+)+)\b/.exec(spec.product || "");
  if (!match) die(`cannot read a LibreOffice version from ${spec.product}`);
  return match[1];
}

function latestWhen(node, dates = []) {
  if (Array.isArray(node)) {
    for (const item of node) latestWhen(item, dates);
  } else if (node && typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      if (key === "when" && typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) dates.push(value);
      else latestWhen(value, dates);
    }
  }
  return dates;
}

function engineRow(id, engine, version, mode, kind) {
  return { id, engine, version, mode, kind, passed: 0, total: 0, percent: "" };
}

if (!excelMatch(0.707106781187, 0.7071067811865476)) die("tolerance self-check failed");
if (excelMatch(true, 1) || excelMatch(false, 0)) die("bool/number self-check failed");
if (!excelMatch("#NUM!", "#NUM!") || !excelMatch(false, false)) die("exact self-check failed");

const cells = loadCells();
const hook = indexResults(baseline.univer.hook.results, "Univer hook");
const native = indexResults(baseline.univer.native.results, "Univer native");
const libre = indexResults(baseline.libreoffice.results, "LibreOffice");
cover(hook, cells, "Univer hook");
cover(native, cells, "Univer native");
cover(libre, cells, "LibreOffice");

const pyFail = indexFailures(baseline.ironcalc_failures, "IronCalc Python");
const jsFail = indexFailures(baseline.ironcalc_js.failures, "IronCalc JS");
if (Object.keys(baseline.ironcalc_js.divergences || {}).length !== 0) {
  die("ironcalc_js.divergences is not empty; this generator assumes JS matches Python on every cell");
}
if (baseline.ironcalc_js.cells_match_python !== cells.length) {
  die(`ironcalc_js.cells_match_python is ${baseline.ironcalc_js.cells_match_python}, cells ${cells.length}`);
}
if (baseline.univer.hook.cells_match_ironcalc !== cells.length) {
  die(`univer hook cells_match_ironcalc is ${baseline.univer.hook.cells_match_ironcalc}, cells ${cells.length}`);
}

for (const cell of cells) {
  const key = `${cell.fixture} ${cell.cell}`;
  const actual = hook.get(key);
  const py = pyFail.get(key);
  const js = jsFail.get(key);
  const misses = !excelMatch(actual, cell.expected);
  if (py === undefined && misses) die(`IronCalc Python has no failure for miss ${key} (${showValue(actual)})`);
  if (py !== undefined && !misses) die(`IronCalc Python lists a match as a failure: ${key}`);
  if (py !== undefined && !valuesEqual(py, actual)) die(`IronCalc Python ${key} ${showValue(py)} != hook ${showValue(actual)}`);
  if (js === undefined && misses) die(`IronCalc JS has no failure for miss ${key}`);
  if (js !== undefined && !misses) die(`IronCalc JS lists a match as a failure: ${key}`);
  if (js !== undefined && !valuesEqual(js, actual)) die(`IronCalc JS ${key} ${showValue(js)} != hook ${showValue(actual)}`);
}

const sheetjsFailures = {
  read: indexFailures(baseline.sheetjs_read_failures, "SheetJS read"),
  write: indexFailures(baseline.sheetjs_write_failures, "SheetJS write"),
};
for (const [mode, found] of Object.entries(sheetjsFailures)) {
  for (const key of found.keys()) {
    if (!cells.some((cell) => `${cell.fixture} ${cell.cell}` === key)) die(`SheetJS ${mode} names an unknown cell ${key}`);
  }
}

const univerVer = univerVersion(baseline.univer.packages);
const loVer = libreofficeVersion(baseline.libreoffice);
if (baseline.ironcalc_version !== baseline.ironcalc_js.version) {
  die("IronCalc Python and JS versions differ");
}

const engines = [
  engineRow("sheetjs-read", "SheetJS CE", baseline.sheetjs_version, "read", "io"),
  engineRow("sheetjs-write", "SheetJS CE", baseline.sheetjs_version, "write", "io"),
  engineRow("univer-native", "Univer", univerVer, "native", "calc"),
  engineRow("libreoffice", "LibreOffice Calc", loVer, "recalc", "calc"),
  engineRow("ironcalc-js", "IronCalc", baseline.ironcalc_js.version, "JavaScript", "calc"),
  engineRow("ironcalc-py", "IronCalc", baseline.ironcalc_version, "Python", "calc"),
  engineRow("univer-hook", "Univer", univerVer, "hook", "calc"),
];

const recorded = {
  "ironcalc-py": [baseline.cells_pass, baseline.cells],
  "ironcalc-js": [baseline.ironcalc_js.cells_match_excel, baseline.ironcalc_js.cells],
  "univer-hook": [baseline.univer.hook.cells_match_excel, baseline.univer.hook.cells],
  "univer-native": [baseline.univer.native.cells_match_excel, baseline.univer.native.cells],
  "libreoffice": [baseline.libreoffice.cells_match_excel, baseline.libreoffice.cells],
};

const outCells = cells.map((cell) => {
  const key = `${cell.fixture} ${cell.cell}`;
  const enginesForCell = {};
  for (const engine of engines) {
    let value;
    let label = null;
    if (engine.kind === "io") {
      const mode = engine.id === "sheetjs-read" ? "read" : "write";
      const failure = sheetjsFailures[mode].get(key);
      const pass = failure === undefined;
      if (pass) label = "preserved";
      else value = failure;
      enginesForCell[engine.id] = { pass, value: pass ? null : value, label };
      if (pass) engine.passed += 1;
    } else {
      const source = engine.id === "univer-native" ? native : engine.id === "libreoffice" ? libre : hook;
      value = source.get(key);
      const pass = excelMatch(value, cell.expected);
      enginesForCell[engine.id] = { pass, value, label: null };
      if (pass) engine.passed += 1;
    }
    engine.total += 1;
  }
  return {
    fixture: cell.fixture,
    cell: cell.cell,
    formula: cell.formula,
    excel: cell.expected,
    source: cell.source,
    engines: enginesForCell,
  };
});

for (const engine of engines) {
  engine.percent = percent(engine.passed, engine.total);
  const expect = recorded[engine.id];
  if (expect && (engine.passed !== expect[0] || engine.total !== expect[1])) {
    die(`${engine.id} counted ${engine.passed}/${engine.total}, baseline ${expect[0]}/${expect[1]}`);
  }
}

engines.sort((a, b) => {
  const rate = b.passed / b.total - a.passed / a.total;
  if (rate !== 0) return rate;
  if (b.passed !== a.passed) return b.passed - a.passed;
  if (a.total !== b.total) return a.total - b.total;
  const byEngine = a.engine.toLowerCase().localeCompare(b.engine.toLowerCase());
  if (byEngine !== 0) return byEngine;
  return a.mode.toLowerCase().localeCompare(b.mode.toLowerCase());
});

const dates = latestWhen(fixtures);
if (dates.length === 0) die("fixtures.json has no when date");
const updated = dates.sort().at(-1);

const samplesDir = join(siteRoot, "public", "samples");
mkdirSync(samplesDir, { recursive: true });
for (const file of SAMPLES) {
  copyFileSync(join(repoRoot, "fixtures", file), join(samplesDir, file));
}

const ironPkg = JSON.parse(readFileSync(join(siteRoot, "node_modules", "@ironcalc", "wasm", "package.json"), "utf8"));
const xlsxPkg = JSON.parse(readFileSync(join(siteRoot, "node_modules", "xlsx", "package.json"), "utf8"));
if (ironPkg.version !== "0.9.1") die(`@ironcalc/wasm is ${ironPkg.version}, expected 0.9.1`);
if (xlsxPkg.version !== "0.20.3") die(`xlsx is ${xlsxPkg.version}, expected 0.20.3`);
if (ironPkg.license !== "MIT OR Apache-2.0") die(`@ironcalc/wasm license is ${ironPkg.license}`);

const payload = {
  generatedFrom: ["suite/baseline.json", "fixtures/fixtures.json"],
  tolerance: "1e-12",
  cellCount: cells.length,
  fixtureCount: fixtures.length,
  updated,
  engines,
  cells: outCells,
  samples: SAMPLES.map((file) => ({ file, id: file.replace(/\.xlsx$/, "") })),
  runtime: { ironcalc: ironPkg.version, sheetjs: xlsxPkg.version },
};

if (payload.fixtureCount !== baseline.fixtures) {
  die(`fixture count ${payload.fixtureCount} != baseline.fixtures ${baseline.fixtures}`);
}

const outDir = join(siteRoot, "src", "generated");
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, "scoreboard.json"), `${JSON.stringify(payload)}\n`);
console.log(
  `scoreboard: ${payload.cellCount} cells, ${engines.map((engine) => `${engine.engine} ${engine.mode} ${engine.passed}/${engine.total}`).join("; ")}`,
);
