"""Write SCOREBOARD.md from the committed baseline and fixture metadata.

Scores are counted from suite/baseline.json. Expected values and their
citations come from fixtures/fixtures.json. Nothing in the table is typed
by hand. The Excel match is the one suite/run.py uses (EXCEL_REL_TOL):
strings, booleans, and errors are exact, and numbers use that relative
tolerance. Regenerating does not read the clock. The last-updated line is
the latest `when` date stored on a fixture source. The Provenance section
counts cells by source type from each case's source string.
"""
import ast
import difflib
import json
import math
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASELINE_PATH = os.path.join(ROOT, "suite", "baseline.json")
FIXTURES_PATH = os.path.join(ROOT, "fixtures", "fixtures.json")
RUN_PATH = os.path.join(ROOT, "suite", "run.py")
SCOREBOARD_PATH = os.path.join(ROOT, "SCOREBOARD.md")
WHEN_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def die(message):
    sys.stderr.write(f"{message}\n")
    raise SystemExit(1)


def load_json(path):
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def excel_rel_tol():
    """The tolerance suite/run.py assigns to EXCEL_REL_TOL. One source."""
    tree = ast.parse(open(RUN_PATH, encoding="utf-8").read(), filename=RUN_PATH)
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "EXCEL_REL_TOL":
                try:
                    value = ast.literal_eval(node.value)
                except (ValueError, SyntaxError) as exc:
                    die(f"EXCEL_REL_TOL in suite/run.py is not a literal: {exc}")
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    die(f"EXCEL_REL_TOL in suite/run.py is {value!r}")
                return float(value)
    die("suite/run.py does not assign EXCEL_REL_TOL")


def excel_match(got, expected, rel_tol):
    """Same rule as suite/run.py excel_match. Bool is not a number."""
    if isinstance(expected, bool) or isinstance(got, bool):
        return isinstance(got, bool) and isinstance(expected, bool) and got is expected
    if isinstance(expected, str) or isinstance(got, str):
        return got == expected
    if isinstance(expected, (int, float)) and isinstance(got, (int, float)):
        return math.isclose(got, expected, rel_tol=rel_tol, abs_tol=0.0)
    return False


def show(value):
    """Render a JSON value. Errors stay bare; strings keep their quotes."""
    if isinstance(value, str) and (
        (value.startswith("#") and value.endswith(("!", "?"))) or value == "#N/A"
    ):
        return value
    return json.dumps(value, ensure_ascii=False)


def load_cells(fixtures):
    cells = []
    seen = set()
    for fixture in fixtures:
        fixture_id = fixture["id"]
        for case in fixture["cases"]:
            if len(case) != 5:
                die(f"{fixture_id} case does not have sheet, cell, formula, expected, source")
            sheet, addr, _formula, expected, source = case
            if not isinstance(source, str) or source == "":
                die(f"{fixture_id} {sheet}!{addr} has no source")
            cell = f"{sheet}!{addr}"
            key = (fixture_id, cell)
            if key in seen:
                die(f"duplicate fixture cell {fixture_id} {cell}")
            seen.add(key)
            cells.append({
                "id": fixture_id,
                "cell": cell,
                "expected": expected,
                "source": source,
            })
    if not cells:
        die("fixtures/fixtures.json has no cells")
    return cells


def index_map(mapping, label):
    found = {}
    if not isinstance(mapping, dict):
        die(f"{label} is not an object")
    for fixture_id, cells in mapping.items():
        if not isinstance(cells, dict):
            die(f"{label} {fixture_id} is not an object")
        for cell, actual in cells.items():
            key = (fixture_id, cell)
            if key in found:
                die(f"{label} repeats {fixture_id} {cell}")
            found[key] = actual
    return found


def misses_from_results(cells, results, label, rel_tol):
    """Full recorded results. A miss is a cell outside the Excel match."""
    found = index_map(results, label)
    wanted = {(cell["id"], cell["cell"]) for cell in cells}
    missing = sorted(wanted - set(found))
    extra = sorted(set(found) - wanted)
    if missing or extra:
        die(f"{label} does not cover the fixture cells (missing {missing[:3]}, extra {extra[:3]})")
    misses = []
    for cell in cells:
        actual = found[(cell["id"], cell["cell"])]
        if not excel_match(actual, cell["expected"], rel_tol):
            misses.append({**cell, "actual": actual})
    return misses


def misses_from_failures(cells, failures, label, rel_tol):
    """A failure map lists misses only. Every entry must miss Excel."""
    found = index_map(failures, label)
    wanted = {(cell["id"], cell["cell"]) for cell in cells}
    extra = sorted(set(found) - wanted)
    if extra:
        die(f"{label} names cells that are not in the fixtures: {extra[:3]}")
    by_key = {(cell["id"], cell["cell"]): cell for cell in cells}
    misses = []
    for cell in cells:
        key = (cell["id"], cell["cell"])
        if key not in found:
            continue
        actual = found[key]
        if excel_match(actual, cell["expected"], rel_tol):
            die(
                f"{label} records {cell['id']} {cell['cell']} as a miss, "
                f"but {show(actual)} matches Excel {show(cell['expected'])}"
            )
        misses.append({**by_key[key], "actual": actual})
    return misses


def require_count(label, passed, total, recorded_passed, recorded_total):
    if recorded_total != total:
        die(f"{label} cells: baseline {recorded_total}, fixtures {total}")
    if recorded_passed != passed:
        die(f"{label} passing: baseline {recorded_passed}, counted {passed}")


def row(engine, version, mode, misses, total, empty=None):
    if not isinstance(version, str) or version == "":
        die(f"{engine} {mode} has no version")
    return {
        "engine": engine,
        "version": version,
        "mode": mode,
        "passed": total - len(misses),
        "total": total,
        "misses": misses,
        "empty": empty,
    }


def univer_version(packages):
    versions = sorted({
        version for name, version in packages.items() if name.startswith("@univerjs/")
    })
    if versions == []:
        die("baseline univer.packages has no @univerjs version")
    if len(versions) != 1:
        die(f"Univer packages do not share one version: {versions}")
    return versions[0]


def libreoffice_version(spec):
    product = spec.get("product") or ""
    match = re.match(r"LibreOffice (\d+(?:\.\d+)+)\b", product)
    if not match:
        die(f"cannot read a LibreOffice version from {product!r}")
    return match.group(1)


def sheetjs_rows(baseline, cells, rel_tol):
    known = {"sheetjs_version", "sheetjs_read_failures", "sheetjs_write_failures"}
    unknown = sorted(key for key in baseline if "sheetjs" in key.lower() and key not in known)
    if unknown:
        die(f"unrecognized SheetJS keys in the baseline: {unknown}")
    version = baseline["sheetjs_version"]
    total = baseline["cells"]
    if total != len(cells):
        die(f"baseline cells {total} != fixture cells {len(cells)}")
    rows = []
    for mode in ("read", "write"):
        failures = baseline.get(f"sheetjs_{mode}_failures", {})
        misses = misses_from_failures(cells, failures, f"SheetJS {mode}", rel_tol)
        empty = (
            f"No misses. The baseline records no SheetJS {mode} failures, "
            "so this mode passes every formula cell."
        )
        rows.append(row("SheetJS CE", version, mode, misses, total, empty=empty))
    return rows


def build_rows(baseline, cells, rel_tol):
    total = len(cells)
    iron_misses = misses_from_failures(
        cells, baseline["ironcalc_failures"], "IronCalc", rel_tol
    )
    iron = row("IronCalc", baseline["ironcalc_version"], "Python", iron_misses, total)
    require_count("IronCalc Python", iron["passed"], total, baseline["cells_pass"], baseline["cells"])

    js = baseline["ironcalc_js"]
    js_misses = misses_from_failures(cells, js["failures"], "IronCalc JS", rel_tol)
    javascript = row("IronCalc", js["version"], "JavaScript", js_misses, total)
    require_count(
        "IronCalc JavaScript",
        javascript["passed"],
        total,
        js["cells_match_excel"],
        js["cells"],
    )

    packages = baseline["univer"]["packages"]
    version = univer_version(packages)
    hook_spec = baseline["univer"]["hook"]
    native_spec = baseline["univer"]["native"]
    hook_misses = misses_from_results(cells, hook_spec["results"], "Univer hook", rel_tol)
    native_misses = misses_from_results(cells, native_spec["results"], "Univer native", rel_tol)
    hook = row("Univer", version, "hook", hook_misses, total)
    native = row("Univer", version, "native", native_misses, total)
    require_count(
        "Univer hook", hook["passed"], total, hook_spec["cells_match_excel"], hook_spec["cells"]
    )
    require_count(
        "Univer native",
        native["passed"],
        total,
        native_spec["cells_match_excel"],
        native_spec["cells"],
    )

    lo = baseline["libreoffice"]
    lo_misses = misses_from_results(cells, lo["results"], "LibreOffice", rel_tol)
    calc = row("LibreOffice Calc", libreoffice_version(lo), "recalc", lo_misses, total)
    require_count(
        "LibreOffice", calc["passed"], total, lo["cells_match_excel"], lo["cells"]
    )

    rows = [iron, javascript, hook, native, calc, *sheetjs_rows(baseline, cells, rel_tol)]
    rows.sort(key=lambda item: (
        -item["passed"] / item["total"],
        -item["passed"],
        item["total"],
        item["engine"].casefold(),
        item["mode"].casefold(),
    ))
    return rows


def latest_when(fixtures):
    dates = []

    def walk(node):
        if isinstance(node, dict):
            for key, value in node.items():
                if key == "when" and isinstance(value, str) and WHEN_RE.fullmatch(value):
                    dates.append(value)
                else:
                    walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(fixtures)
    if not dates:
        die("fixtures/fixtures.json has no `when` date")
    return max(dates)


def provenance(fixtures):
    """One line. Cell counts are the Provenance section."""
    if not fixtures:
        die("fixtures/fixtures.json has no fixtures")
    return (
        "Expected values come from recorded desktop Excel runs, "
        "plus the remaining cited sources. "
        "HyperFormula's differences page is a reference, not the source of the values. "
        "Each miss cites the source stored on that cell, and IronCalc's misses are grouped in "
        "[docs/ironcalc-triage.md](docs/ironcalc-triage.md)."
    )


# Order is the order printed. Classification reads only the text before
# "; previously:", so a kept citation does not change the cell's type.
SOURCE_TYPES = (
    "Adam's desktop Excel run",
    "third-party report of an Excel run",
    "issue-stated value",
    "Microsoft docs",
)


def classify_source(source):
    """Map one case source string to a SOURCE_TYPES label."""
    primary = source.split("; previously:", 1)[0].strip()
    if primary.startswith("primary: Adam Marsh"):
        return SOURCE_TYPES[0]
    if primary.startswith("MS "):
        return SOURCE_TYPES[3]
    lowered = primary.casefold()
    names_run = "excel-saved" in lowered or (
        re.search(r"\bexcel\b", lowered) is not None
        and re.search(r"\b(?:ran|run|copied from)\b", lowered) is not None
    )
    if names_run:
        return SOURCE_TYPES[1]
    return SOURCE_TYPES[2]


def render_provenance(cells):
    counts = {name: 0 for name in SOURCE_TYPES}
    for cell in cells:
        counts[classify_source(cell["source"])] += 1
    if sum(counts.values()) != len(cells):
        die("provenance counts do not cover every fixture cell")
    lines = [
        "## Provenance",
        "",
        "Cell counts by source type. "
        "Classification uses the text before `; previously:`: "
        "a `primary: Adam Marsh` prefix is Adam's desktop Excel run; "
        "a source starting with `MS ` is Microsoft docs; "
        "a source that names an Excel run or an Excel-saved workbook is a third-party report of an Excel run; "
        "anything else is an issue-stated value.",
        "",
        "| Source type | Cells |",
        "|---|---|",
    ]
    for name in SOURCE_TYPES:
        lines.append(f"| {name} | {counts[name]} |")
    lines.append("")
    return lines


def percent(passed, total):
    if passed == total:
        return "100%"
    tenths = (passed * 1000 + total // 2) // total
    return f"{tenths // 10}.{tenths % 10}%"


def plain(text):
    return text.replace("|", "\\|").replace("\n", " ")


def code(value):
    return "`" + show(value).replace("`", "'") + "`"


def miss_intro(count, empty):
    if count == 0:
        return empty or "No misses."
    if count == 1:
        return "1 cell misses."
    return f"{count} cells miss."


def render_misses(rows):
    lines = ["## Misses", ""]
    for item in rows:
        lines.append(f"### {item['engine']} {item['version']} {item['mode']}")
        lines.append("")
        lines.append(miss_intro(len(item["misses"]), item["empty"]))
        lines.append("")
        grouped = []
        for miss in item["misses"]:
            if not grouped or grouped[-1][0] != miss["id"]:
                grouped.append((miss["id"], []))
            grouped[-1][1].append(miss)
        for fixture_id, misses in grouped:
            lines.append(f"#### {fixture_id}")
            lines.append("")
            lines.append("| Cell | Expected | Actual | Source |")
            lines.append("|---|---|---|---|")
            for miss in misses:
                lines.append(
                    "| "
                    + " | ".join((
                        plain(miss["cell"]),
                        code(miss["expected"]),
                        code(miss["actual"]),
                        plain(miss["source"]),
                    ))
                    + " |"
                )
            lines.append("")
    return lines


def format_tol(value):
    return format(value, ".0e").replace("e-0", "e-").replace("e+0", "e+")


def render():
    baseline = load_json(BASELINE_PATH)
    fixtures = load_json(FIXTURES_PATH)
    rel_tol = excel_rel_tol()
    cells = load_cells(fixtures)
    rows = build_rows(baseline, cells, rel_tol)
    updated = latest_when(fixtures)
    lines = [
        "# Scoreboard",
        "",
        f"Last updated: {updated} (latest `when` in fixtures/fixtures.json).",
        "",
        provenance(fixtures),
        "",
        "Generated from [`suite/baseline.json`](suite/baseline.json) and "
        "[`fixtures/fixtures.json`](fixtures/fixtures.json) by `make scoreboard`.",
        "",
        "| Engine | Version | Mode | Passing | Percent |",
        "|---|---|---|---|---|",
    ]
    for item in rows:
        passing = f"{item['passed']}/{item['total']}"
        lines.append(
            "| "
            + " | ".join((
                plain(item["engine"]),
                plain(item["version"]),
                plain(item["mode"]),
                passing,
                percent(item["passed"], item["total"]),
            ))
            + " |"
        )
    lines.append("")
    lines.append(
        "A cell passes when the recorded value matches Excel. "
        f"Strings, booleans, and errors are exact. Numbers use a relative tolerance of {format_tol(rel_tol)}. "
        "SheetJS read and write pass when the formula string and the cached value are preserved."
    )
    lines.append("")
    lines.append(
        "The Univer hook is IronCalc running inside Univer. "
        "SheetJS CE does not calculate, so its row is a read/write (I/O) check."
    )
    lines.append("")
    lines.extend(render_provenance(cells))
    lines.extend(render_misses(rows))
    text = "\n".join(lines).rstrip() + "\n"
    return text


def check(text):
    if not os.path.exists(SCOREBOARD_PATH):
        sys.stderr.write("SCOREBOARD.md is missing. Run make scoreboard.\n")
        return 1
    with open(SCOREBOARD_PATH, encoding="utf-8") as handle:
        current = handle.read()
    if current == text:
        print("SCOREBOARD.md is up to date.")
        return 0
    sys.stderr.write("SCOREBOARD.md is stale. Run make scoreboard.\n")
    diff = "".join(difflib.unified_diff(
        current.splitlines(keepends=True),
        text.splitlines(keepends=True),
        fromfile="SCOREBOARD.md",
        tofile="SCOREBOARD.md (generated)",
    ))
    sys.stderr.write(diff)
    if diff and not diff.endswith("\n"):
        sys.stderr.write("\n")
    return 1


def main(argv):
    if argv not in ([], ["--check"]):
        sys.stderr.write("usage: scoreboard.py [--check]\n")
        return 2
    text = render()
    if argv == ["--check"]:
        return check(text)
    with open(SCOREBOARD_PATH, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)
    print(f"wrote {os.path.relpath(SCOREBOARD_PATH, ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
