// IronCalc's Node.js build calculates the same fixtures as the Python bindings.
// @ironcalc/nodejs loads xlsx directly, so the wasm package is not used.
const fs = require("fs");
const path = require("path");
const { Model } = require("@ironcalc/nodejs");

const outDir = process.argv[2];
if (!outDir) {
  console.error("usage: ironcalc_js.js <out-dir>");
  process.exit(2);
}
fs.mkdirSync(outDir, { recursive: true });

const root = path.resolve(__dirname, "..");
const pkg = require(path.join(root, "node_modules", "@ironcalc", "nodejs", "package.json"));
const fixtures = JSON.parse(fs.readFileSync(path.join(root, "fixtures", "fixtures.json"), "utf8"));
const cases = [];

function errorText(exc) {
  return exc instanceof Error ? `${exc.name}: ${exc.message}` : String(exc);
}

for (const fx of fixtures) {
  const abs = path.join(root, "fixtures", fx.file);
  let model = null;
  let loadError = null;
  try {
    // locale, timezone, language match ironcalc.load_from_xlsx(path, "en", "UTC").
    model = Model.fromXlsx(abs, "en", "UTC", "en");
    model.evaluate();
  } catch (exc) {
    loadError = errorText(exc);
  }
  for (const [sheet, cell] of fx.cases) {
    const ref = `${sheet}!${cell}`;
    let got = null;
    let error = loadError;
    if (!loadError) {
      try {
        got = model.getCellValueByRef(ref);
        if (got === undefined) got = null;
      } catch (exc) {
        got = null;
        error = errorText(exc);
      }
    }
    cases.push({ file: fx.file, cell: ref, got, error });
  }
}

const report = { package: pkg.name, version: pkg.version, cases };
fs.writeFileSync(path.join(outDir, "read.json"), JSON.stringify(report));
