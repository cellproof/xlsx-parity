import init, { Model } from "@ironcalc/wasm";
import scoreboard from "./generated/scoreboard.json";
import { showValue } from "./compare.js";
import { LOADER_SUMMARY, loaderDetail, scanWorkbook } from "./scan.js";

const IRON_VERSION = scoreboard.runtime.ironcalc;

const engineStatus = document.querySelector("#engine-status");
const drop = document.querySelector("#drop");
const fileInput = document.querySelector("#file");
const sampleButtons = document.querySelector("#sample-buttons");
const scanOut = document.querySelector("#scan-out");
const scoreRoot = document.querySelector("#scoreboard-root");

let wasmReady = false;
let scanning = false;

function text(tag, value, attrs) {
  const node = document.createElement(tag);
  if (attrs) {
    for (const [key, attr] of Object.entries(attrs)) {
      if (key === "class") node.className = attr;
      else node.setAttribute(key, attr);
    }
  }
  if (value != null) node.textContent = value;
  return node;
}

function setEngineStatus(state, message) {
  engineStatus.dataset.engine = state;
  engineStatus.dataset.ironcalc = IRON_VERSION;
  engineStatus.dataset.nativeXlsx = "false";
  engineStatus.textContent = message;
  for (const button of sampleButtons.querySelectorAll("button")) {
    button.disabled = state !== "ready" || scanning;
  }
  fileInput.disabled = state !== "ready" || scanning;
}

function renderSamples() {
  for (const sample of scoreboard.samples) {
    const button = text("button", sample.id, { type: "button", "data-sample": sample.file });
    button.disabled = true;
    button.addEventListener("click", () => loadSample(sample));
    sampleButtons.append(button);
  }
}

async function loadSample(sample) {
  try {
    const url = `${import.meta.env.BASE_URL}samples/${sample.file}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Sample ${sample.file} returned ${response.status}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    await runScan(bytes, sample.file);
  } catch (err) {
    scanOut.hidden = false;
    scanOut.replaceChildren(text("p", err instanceof Error ? err.message : String(err), { class: "summary", "data-scan-status": "error" }));
  }
}

async function runScan(bytes, fileName) {
  if (!wasmReady || scanning) return;
  scanning = true;
  setEngineStatus("ready", `IronCalc wasm: ready (@ironcalc/wasm ${IRON_VERSION}). Scanning ${fileName}.`);
  scanOut.hidden = false;
  scanOut.replaceChildren(text("p", `Scanning ${fileName}…`, { class: "meta" }));
  await new Promise((resolve) => setTimeout(resolve, 0));
  let model = null;
  try {
    model = new Model(fileName, "en", "UTC", "en");
    const result = scanWorkbook(model, bytes, fileName);
    result.ironcalc = IRON_VERSION;
    result.note = loaderDetail(result.sheetjs, IRON_VERSION);
    renderScan(result);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    scanOut.replaceChildren(text("p", message, { class: "summary", "data-scan-status": "error" }));
  } finally {
    if (model) {
      try { model.free(); } catch { /* the model is already released */ }
    }
    scanning = false;
    setEngineStatus("ready", `IronCalc wasm: ready (@ironcalc/wasm ${IRON_VERSION}). No native .xlsx import.`);
  }
}

function renderScan(result) {
  const { summary } = result;
  const banner = text("div", null, {
    class: "summary",
    id: "scan-summary",
    "data-scan-status": "done",
    "data-loader": result.loader,
    "data-native-xlsx": "false",
    "data-match": String(summary.match),
    "data-total": String(summary.formulaCells),
    "data-mismatch": String(summary.mismatch),
    "data-uncached": String(summary.uncached),
    "data-error": String(summary.error),
    "data-file": result.file,
  });
  banner.append(
    text("strong", `${summary.match} / ${summary.formulaCells} formula cells match Excel's cached values`),
    text("p", `${summary.mismatch} ${summary.mismatch === 1 ? "mismatch" : "mismatches"} · ${summary.uncached} no cached Excel value · ${summary.error} ${summary.error === 1 ? "error" : "errors"}`, { class: "counts" }),
    text("p", LOADER_SUMMARY, { class: "note" }),
    text("p", "Cells with no cached Excel value are counted separately. Excel writes a cached value by default; a workbook from a script may not.", { class: "note" }),
  );

  const details = document.createElement("details");
  details.append(text("summary", "How this load differs from the suite"));
  details.append(text("p", result.note));
  details.append(text("p", "Numbers use a relative tolerance of 1e-12 and an absolute tolerance of 0. Strings, booleans, and errors are exact. Same rule as suite/run.py excel_match. IronCalc numbers are read back through a 16-digit scientific format because getFormattedCellValue is the value this package exposes."));
  banner.append(details);
  if (result.warnings.length) {
    const list = text("ul", null, { class: "warn" });
    for (const warning of result.warnings) list.append(text("li", warning));
    banner.append(list);
  }

  const toolbar = text("div", null, { class: "toolbar" });
  const mismatchOnly = document.createElement("input");
  mismatchOnly.type = "checkbox";
  mismatchOnly.checked = true;
  mismatchOnly.id = "mismatch-only";
  const mismatchLabel = text("label", null);
  mismatchLabel.append(mismatchOnly, document.createTextNode(" Mismatches only"));
  const allHint = text("span", "Uncheck for all formula cells", { class: "meta" });
  const search = document.createElement("input");
  search.type = "search";
  search.placeholder = "Sheet, cell, formula, value";
  search.setAttribute("aria-label", "Filter formula cells");
  const sheetSelect = document.createElement("select");
  sheetSelect.setAttribute("aria-label", "Sheet");
  sheetSelect.append(text("option", "All sheets", { value: "" }));
  for (const name of [...new Set(result.cells.map((row) => row.sheet))]) {
    sheetSelect.append(text("option", name, { value: name }));
  }
  const copy = text("button", "Copy Markdown report", { type: "button" });
  const download = text("button", "Download JSON", { type: "button" });
  toolbar.append(mismatchLabel, allHint, sheetSelect, search, copy, download);

  const shown = text("p", "", { class: "meta", id: "scan-shown" });
  const wrap = text("div", null, { class: "table-wrap" });
  const table = document.createElement("table");
  table.id = "scan-table";
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const label of ["Sheet", "Cell", "Formula", "Excel (cached)", `IronCalc ${IRON_VERSION}`, "Status"]) {
    headRow.append(text("th", label));
  }
  head.append(headRow);
  const body = document.createElement("tbody");
  table.append(head, body);
  wrap.append(table);

  function excelText(row) {
    return row.excelCached ? showValue(row.excel) : "";
  }
  function ironText(row) {
    return row.error ? row.error : showValue(row.ironcalc);
  }

  function paint() {
    const query = search.value.trim().toLowerCase();
    const sheet = sheetSelect.value;
    const onlyMisses = mismatchOnly.checked;
    const visible = result.cells.filter((row) => {
      if (onlyMisses && row.status === "match") return false;
      if (sheet && row.sheet !== sheet) return false;
      if (!query) return true;
      const hay = [row.sheet, row.cell, row.formula, excelText(row), ironText(row), row.status].join(" ").toLowerCase();
      return hay.includes(query);
    });
    body.replaceChildren();
    const capped = visible.slice(0, 1000);
    for (const row of capped) {
      const tr = document.createElement("tr");
      tr.className = row.status;
      tr.dataset.sheet = row.sheet;
      tr.dataset.cell = `${row.sheet}!${row.cell}`;
      tr.dataset.status = row.status;
      tr.append(
        text("td", row.sheet),
        text("td", row.cell),
        text("td", row.formula, { class: "formula" }),
        text("td", row.excelCached ? excelText(row) : "(no cached value)", { "data-col": "excel" }),
        text("td", ironText(row), { "data-col": "iron" }),
        text("td", statusLabel(row.status), { class: "status" }),
      );
      body.append(tr);
    }
    if (capped.length === 0) {
      const tr = document.createElement("tr");
      const td = text("td", "No formula cells in this view.", { class: "empty" });
      td.colSpan = 6;
      tr.append(td);
      body.append(tr);
    }
    const extra = visible.length - capped.length;
    shown.textContent = extra > 0
      ? `Showing ${capped.length} of ${visible.length} formula cells in this view. Download JSON for the rest.`
      : `Showing ${visible.length} formula cells.`;
  }

  mismatchOnly.addEventListener("change", paint);
  search.addEventListener("input", paint);
  sheetSelect.addEventListener("change", paint);
  copy.addEventListener("click", async () => {
    const markdown = markdownReport(result);
    try {
      await navigator.clipboard.writeText(markdown);
      copy.textContent = "Copied";
    } catch {
      copy.textContent = "Copy failed";
    }
    setTimeout(() => { copy.textContent = "Copy Markdown report"; }, 1500);
  });
  download.addEventListener("click", () => {
    const payload = {
      file: result.file,
      sheetjs: result.sheetjs,
      ironcalc: result.ironcalc,
      loader: result.loader,
      nativeXlsxImport: false,
      tolerance: { numericRelative: 1e-12, absolute: 0, strings: "exact", booleans: "exact", errors: "exact" },
      summary: result.summary,
      warnings: result.warnings,
      note: result.note,
      cells: result.cells,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2) + "\n"], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${result.file.replace(/\.xlsx$/i, "")}-ironcalc.json`;
    link.click();
    URL.revokeObjectURL(url);
  });

  scanOut.replaceChildren(banner, toolbar, shown, wrap);
  paint();
}

function missLabel(count) {
  if (count === 1) return "1 miss";
  return `${count} misses`;
}

function statusLabel(status) {
  if (status === "uncached") return "No cached Excel value";
  if (status === "mismatch") return "Mismatch";
  if (status === "match") return "Match";
  return "Error";
}

function markdownReport(result) {
  const { summary } = result;
  const lines = [
    `# xlsx-parity scan: ${result.file}`,
    "",
    LOADER_SUMMARY,
    "",
    `SheetJS CE ${result.sheetjs}. @ironcalc/wasm ${result.ironcalc}.`,
    "",
    `${summary.match} / ${summary.formulaCells} formula cells match Excel's cached values.`,
    "",
    `Mismatch: ${summary.mismatch}. No cached Excel value: ${summary.uncached}. Errors: ${summary.error}.`,
    "",
    "Numbers use a relative tolerance of 1e-12 and an absolute tolerance of 0. Strings, booleans, and errors are exact.",
    "",
    result.note,
    "",
  ];
  if (result.warnings.length) {
    lines.push("Notes:", "");
    for (const warning of result.warnings) lines.push(`- ${warning.replace(/\n/g, " ")}`);
    lines.push("");
  }
  lines.push(`| Sheet | Cell | Formula | Excel (cached) | IronCalc ${result.ironcalc} | Status |`);
  lines.push("|---|---|---|---|---|---|");
  for (const row of result.cells) {
    const excel = row.excelCached ? showValue(row.excel) : "(no cached value)";
    const iron = row.error ? row.error : showValue(row.ironcalc);
    lines.push(`| ${md(row.sheet)} | ${md(row.cell)} | ${md(row.formula)} | ${md(excel)} | ${md(iron)} | ${statusLabel(row.status)} |`);
  }
  lines.push("");
  return lines.join("\n");
}

function md(value) {
  return String(value).replace(/\|/g, "\\|").replace(/\n/g, " ");
}

function renderScoreboard() {
  const intro = text("p", null, { class: "meta", id: "scoreboard-meta", "data-cells": String(scoreboard.cellCount), "data-fixtures": String(scoreboard.fixtureCount) });
  intro.append(
    document.createTextNode(`${scoreboard.cellCount} formula cells, ${scoreboard.fixtureCount} fixtures. `),
    document.createTextNode(`Built from ${scoreboard.generatedFrom.join(" and ")}. Latest source date ${scoreboard.updated}. `),
    document.createTextNode("A cell passes when the recorded value matches Excel. Strings, booleans, and errors are exact. Numbers use relative tolerance 1e-12. IronCalc Python, IronCalc JavaScript, and the Univer hook record the same value on every cell, so those three columns share that recorded value."),
  );

  const cards = text("div", null, { class: "cards" });
  for (const engine of scoreboard.engines) {
    const card = text("article", null, {
      class: "card",
      "data-engine-id": engine.id,
      "data-passed": String(engine.passed),
      "data-total": String(engine.total),
      "data-percent": engine.percent,
    });
    card.append(
      text("div", engine.engine, { class: "who" }),
      text("div", `${engine.version} · ${engine.mode}`, { class: "mode" }),
      text("div", `${engine.passed}/${engine.total}`, { class: "ratio" }),
      text("div", `${engine.percent} · ${missLabel(engine.total - engine.passed)}`, { class: "pct" }),
    );
    cards.append(card);
  }

  const strip = text("div", null, { class: "strip", id: "upstream" });
  strip.append(text("span", "Upstream fixes from this suite"));
  const links = [
    ["https://github.com/ironcalc/IronCalc/pull/1791", "#1791 merged", "COUNTA counts an omitted argument"],
    ["https://github.com/ironcalc/IronCalc/pull/1807", "#1807 merged", "T.INV(0, 1) returns #DIV/0!"],
    ["https://github.com/ironcalc/IronCalc/pull/1815", "#1815 open", "LARGE and SMALL when k is not an integer or a boolean"],
    ["https://github.com/ironcalc/IronCalc/pull/1816", "#1816 open", "SUMIFS criteria on legacy xlsx import"],
    ["https://github.com/ironcalc/IronCalc/pull/1817", "#1817 open", "ranges in the branches of IF"],
  ];
  for (const [href, label, title] of links) {
    const link = text("a", `${label} · ${title}`, { href, title });
    link.target = "_blank";
    link.rel = "noreferrer";
    strip.append(link);
  }

  const toolbar = text("div", null, { class: "toolbar" });
  const fixtureSelect = document.createElement("select");
  fixtureSelect.setAttribute("aria-label", "Fixture");
  fixtureSelect.append(text("option", "All fixtures", { value: "" }));
  for (const id of [...new Set(scoreboard.cells.map((cell) => cell.fixture))]) {
    fixtureSelect.append(text("option", id, { value: id }));
  }
  const engineSelect = document.createElement("select");
  engineSelect.setAttribute("aria-label", "Engine");
  engineSelect.append(text("option", "All engines", { value: "" }));
  for (const engine of scoreboard.engines) {
    engineSelect.append(text("option", `${engine.engine} ${engine.mode}`, { value: engine.id }));
  }
  const resultSelect = document.createElement("select");
  resultSelect.setAttribute("aria-label", "Pass or fail");
  resultSelect.append(text("option", "Pass and fail", { value: "" }));
  resultSelect.append(text("option", "Pass", { value: "pass" }));
  resultSelect.append(text("option", "Fail", { value: "fail" }));
  const search = document.createElement("input");
  search.type = "search";
  search.placeholder = "Fixture, cell, formula";
  search.setAttribute("aria-label", "Filter scoreboard");
  toolbar.append(fixtureSelect, engineSelect, resultSelect, search);

  const shown = text("p", "", { class: "meta" });
  const wrap = text("div", null, { class: "table-wrap" });
  const table = document.createElement("table");
  table.id = "scoreboard-table";
  const head = document.createElement("thead");
  const body = document.createElement("tbody");
  table.append(head, body);
  wrap.append(table);

  function engineCell(entry) {
    const td = document.createElement("td");
    td.className = entry.pass ? "pass" : "fail";
    td.dataset.pass = entry.pass ? "true" : "false";
    if (!entry.pass) td.append(text("span", "fail", { class: "tag" }));
    td.append(document.createTextNode(entry.label || showValue(entry.value)));
    return td;
  }

  function paint() {
    const fixture = fixtureSelect.value;
    const engineId = engineSelect.value;
    const result = resultSelect.value;
    const query = search.value.trim().toLowerCase();
    const columns = engineId ? scoreboard.engines.filter((engine) => engine.id === engineId) : scoreboard.engines;
    head.replaceChildren();
    const headRow = document.createElement("tr");
    for (const label of ["Fixture", "Cell", "Formula", "Excel"]) headRow.append(text("th", label));
    for (const engine of columns) {
      headRow.append(text("th", `${engine.engine} ${engine.mode}`, { title: `${engine.engine} ${engine.version} ${engine.mode}` }));
    }
    head.append(headRow);

    const visible = scoreboard.cells.filter((cell) => {
      if (fixture && cell.fixture !== fixture) return false;
      if (query && !`${cell.fixture} ${cell.cell} ${cell.formula}`.toLowerCase().includes(query)) return false;
      if (!result) return true;
      const considered = columns.map((engine) => cell.engines[engine.id].pass);
      const ok = considered.every(Boolean);
      return result === "pass" ? ok : !ok;
    });
    body.replaceChildren();
    for (const cell of visible) {
      const tr = document.createElement("tr");
      tr.dataset.fixture = cell.fixture;
      tr.dataset.cell = cell.cell;
      tr.append(
        text("td", cell.fixture),
        text("td", cell.cell),
        text("td", cell.formula, { class: "formula", title: cell.source }),
        text("td", showValue(cell.excel), { "data-col": "excel" }),
      );
      for (const engine of columns) tr.append(engineCell(cell.engines[engine.id]));
      body.append(tr);
    }
    shown.textContent = `Showing ${visible.length} of ${scoreboard.cellCount} cells.`;
  }

  for (const control of [fixtureSelect, engineSelect, resultSelect]) control.addEventListener("change", paint);
  search.addEventListener("input", paint);
  scoreRoot.replaceChildren(intro, cards, strip, toolbar, shown, wrap);
  paint();
}

function takeFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    scanOut.hidden = false;
    scanOut.replaceChildren(text("p", `${file.name} is not a .xlsx file.`, { class: "summary", "data-scan-status": "error" }));
    return;
  }
  file.arrayBuffer().then((buffer) => runScan(new Uint8Array(buffer), file.name)).catch((err) => {
    scanOut.hidden = false;
    scanOut.replaceChildren(text("p", err instanceof Error ? err.message : String(err), { class: "summary", "data-scan-status": "error" }));
  });
}

drop.addEventListener("dragover", (event) => {
  event.preventDefault();
  drop.classList.add("over");
});
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (event) => {
  event.preventDefault();
  drop.classList.remove("over");
  takeFile(event.dataTransfer.files && event.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
  takeFile(fileInput.files && fileInput.files[0]);
  fileInput.value = "";
});

renderSamples();
renderScoreboard();
init()
  .then(() => {
    wasmReady = true;
    setEngineStatus("ready", `IronCalc wasm: ready (@ironcalc/wasm ${IRON_VERSION}). No native .xlsx import.`);
  })
  .catch((err) => {
    const message = err instanceof Error ? err.message : String(err);
    setEngineStatus("error", `IronCalc wasm failed to load: ${message}`);
  });
