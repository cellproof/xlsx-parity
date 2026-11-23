// SheetJS CE does not calculate. Check that formula strings and cached values
// are read intact, then write a round-trip copy for the XML comparison in run.py.
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const outDir = process.argv[2];
if (!outDir) {
  console.error("usage: sheetjs_read.js <out-dir>");
  process.exit(2);
}
fs.mkdirSync(outDir, { recursive: true });

const root = path.resolve(__dirname, "..");
const fixtures = JSON.parse(fs.readFileSync(path.join(root, "fixtures", "fixtures.json"), "utf8"));
const files = fixtures.map((f) => f.file).concat(["RT-features.xlsx"]);
const report = { version: XLSX.version, cases: [] };

for (const file of files) {
  const abs = path.join(root, "fixtures", file);
  const fx = fixtures.find((f) => f.file === file);
  const wb = XLSX.readFile(abs, {
    cellFormula: true,
    xlfn: true,
    cellNF: true,
    cellStyles: true,
    cellDates: false,
  });
  if (fx) {
    for (const [sn, cell, formula, expected] of fx.cases) {
      const c = wb.Sheets[sn] && wb.Sheets[sn][cell];
      const formulaIntact = !!c && c.f === formula;
      let cachedIntact = false;
      let readCached = null;
      if (c) {
        readCached = c.t === "e" ? c.w : c.v;
        cachedIntact = typeof expected === "string"
          ? (c.t === "e" ? c.w === expected : c.v === expected)
          : c.v === expected;
      }
      report.cases.push({
        file,
        cell: `${sn}!${cell}`,
        formula_intact: formulaIntact,
        cached_intact: cachedIntact,
        read_formula: c ? c.f : null,
        read_cached: readCached,
      });
    }
  }
  XLSX.writeFile(wb, path.join(outDir, path.basename(file)), { bookType: "xlsx", cellStyles: true });
}

fs.writeFileSync(path.join(outDir, "read.json"), JSON.stringify(report));
