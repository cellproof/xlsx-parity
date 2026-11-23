"""Test 1 baseline: IronCalc calculation, SheetJS CE read/write, feature round-trip.

Excel-expected values and their citations live in fixtures/fixtures.json.
Known IronCalc misses, including the JavaScript build, the six cells whose
wrong values are written back as cached values, and each feature kept/lost
result are recorded in suite/baseline.json. Feature results are read from the
written xlsx, one feature and one library at a time. The process exits 0 when
every measured result matches that record, and nonzero when a result changes
— a regression or an unexpected fix.
"""
import json
import os
import re
import subprocess
import sys
from importlib.metadata import version

import openpyxl

import xlsx_xml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES = os.path.join(ROOT, "fixtures")
OUT = os.path.join(ROOT, "out")
WEAKER_ORDER = (
    "F10-sumifs-legacy-whole-column",
    "F07-value-accounting-negatives",
    "F04-text-negative-zero",
)


def same(got, expected):
    if isinstance(expected, str):
        return got == expected
    if isinstance(got, bool) or not isinstance(got, (int, float)):
        return False
    return got == expected


def values_equal(a, b):
    """Cross-engine equality. Bool is not a number; JSON ints still match Python floats."""
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a is b
    if isinstance(a, str) or isinstance(b, str):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b
    return a == b


def show(value):
    if isinstance(value, str):
        if value.startswith("#") and value.endswith("!"):
            return value
        return json.dumps(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def human(formula):
    return re.sub(r"_xl(fn|pm|ws)\.", "", formula)


def recorded_failure(baseline, fixture_id, cell):
    cells = baseline["ironcalc_failures"].get(fixture_id)
    if not cells or cell not in cells:
        return None
    return cells[cell]


def load_fixtures():
    return json.load(open(os.path.join(FIXTURES, "fixtures.json")))


def load_baseline():
    return json.load(open(os.path.join(ROOT, "suite", "baseline.json")))


def run_ironcalc(fixtures):
    import ironcalc

    ver = version("ironcalc")
    rows = []
    ic_dir = os.path.join(OUT, "ironcalc")
    os.makedirs(ic_dir, exist_ok=True)
    for fx in fixtures:
        path = os.path.join(FIXTURES, fx["file"])
        try:
            model = ironcalc.load_from_xlsx(path, "en", "UTC")
            model.evaluate()
            load_error = None
        except BaseException as exc:
            model = None
            load_error = f"{type(exc).__name__}: {exc}"
        for sheet, cell, formula, expected, source in fx["cases"]:
            ref = f"{sheet}!{cell}"
            if load_error:
                got, error = None, load_error
            else:
                try:
                    got = model.get_cell_value_by_ref(ref)
                    error = None
                except Exception as exc:
                    got, error = None, f"{type(exc).__name__}: {exc}"
            rows.append({
                "id": fx["id"],
                "file": fx["file"],
                "sheet": sheet,
                "addr": cell,
                "cell": ref,
                "formula": formula,
                "excel": expected,
                "source": source,
                "got": got,
                "error": error,
            })
        if model is not None:
            dest = os.path.join(ic_dir, fx["file"])
            if os.path.exists(dest):
                os.remove(dest)
            model.save_to_xlsx(dest)
    features_src = os.path.join(FIXTURES, "RT-features.xlsx")
    features_dest = os.path.join(ic_dir, "RT-features.xlsx")
    model = ironcalc.load_from_xlsx(features_src, "en", "UTC")
    model.evaluate()
    if os.path.exists(features_dest):
        os.remove(features_dest)
    model.save_to_xlsx(features_dest)
    return ver, rows


def run_ironcalc_js():
    js_dir = os.path.join(OUT, "ironcalc-js")
    os.makedirs(js_dir, exist_ok=True)
    proc = subprocess.run(
        ["node", os.path.join(ROOT, "suite", "ironcalc_js.js"), js_dir],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        sys.stderr.write(proc.stderr or proc.stdout)
        raise SystemExit(proc.returncode or 1)
    report = json.load(open(os.path.join(js_dir, "read.json")))
    return report["package"], report["version"], {(c["file"], c["cell"]): c for c in report["cases"]}


def js_result(row):
    return row["js_error"] if row["js_error"] else row["js_got"]


def attach_js(rows, js_cases):
    for row in rows:
        js = js_cases.get((row["file"], row["cell"]))
        if js is None:
            row["js_got"] = None
            row["js_error"] = "missing from IronCalc JS report"
        else:
            row["js_got"] = js.get("got")
            row["js_error"] = js.get("error")
        if row["js_error"] or row["error"]:
            row["js_matches_python"] = bool(row["js_error"]) and row["js_error"] == row["error"]
        else:
            row["js_matches_python"] = values_equal(row["js_got"], row["got"])
        row["js_matches_excel"] = row["js_error"] is None and same(row["js_got"], row["excel"])


def recorded_cell(mapping, fixture_id, cell):
    cells = mapping.get(fixture_id)
    if not cells or cell not in cells:
        return False, None
    return True, cells[cell]


def js_counts(rows):
    return {
        "cells_match_excel": sum(bool(row["js_matches_excel"]) for row in rows),
        "cells_match_python": sum(bool(row["js_matches_python"]) for row in rows),
        "cells": len(rows),
        "errors": sum(row["js_error"] is not None for row in rows),
    }


def collect_js_mismatches(rows, baseline, pkg, ver):
    spec = baseline["ironcalc_js"]
    mismatches = []
    if pkg != spec["package"]:
        mismatches.append(f"IronCalc JS package {pkg} != {spec['package']}")
    if ver != spec["version"]:
        mismatches.append(f"IronCalc JS version {ver} != {spec['version']}")
    stats = js_counts(rows)
    for key in ("cells_match_excel", "cells_match_python", "cells", "errors"):
        if stats[key] != spec[key]:
            mismatches.append(f"ironcalc_js {key}: measured {stats[key]}, baseline {spec[key]}")
    known_fail = {(fid, cell) for fid, cells in spec["failures"].items() for cell in cells}
    known_div = {(fid, cell) for fid, cells in spec["divergences"].items() for cell in cells}
    seen = set()
    for row in rows:
        key = (row["id"], row["cell"])
        seen.add(key)
        present, recorded = recorded_cell(spec["failures"], row["id"], row["cell"])
        got = js_result(row)
        if row["js_matches_excel"]:
            if present:
                mismatches.append(
                    f"IronCalc JS {row['id']} {row['cell']}: baseline recorded {show(recorded)}, now matches Excel {show(row['excel'])}"
                )
        elif not present or not values_equal(got, recorded):
            baseline_text = show(recorded) if present else "Excel match"
            got_text = row["js_error"] if row["js_error"] else show(row["js_got"])
            mismatches.append(
                f"IronCalc JS {row['id']} {row['cell']}: baseline {baseline_text}, got {got_text} (Excel {show(row['excel'])})"
            )
        div_present, div_recorded = recorded_cell(spec["divergences"], row["id"], row["cell"])
        if row["js_matches_python"]:
            if div_present:
                py = row["error"] if row["error"] else show(row["got"])
                mismatches.append(
                    f"IronCalc JS {row['id']} {row['cell']}: baseline divergence {show(div_recorded)} now matches Python {py}"
                )
        elif not div_present or not values_equal(got, div_recorded):
            py = row["error"] if row["error"] else show(row["got"])
            js = row["js_error"] if row["js_error"] else show(row["js_got"])
            mismatches.append(
                f"IronCalc JS diverges from Python {row['id']} {row['cell']}: python {py}, js {js}"
            )
    for key in sorted(known_fail - seen):
        mismatches.append(f"ironcalc_js failure {key[0]} {key[1]} is not a fixture cell")
    for key in sorted(known_div - seen):
        mismatches.append(f"ironcalc_js divergence {key[0]} {key[1]} is not a fixture cell")
    return mismatches


def run_sheetjs():
    sj_dir = os.path.join(OUT, "sheetjs")
    os.makedirs(sj_dir, exist_ok=True)
    proc = subprocess.run(
        ["node", os.path.join(ROOT, "suite", "sheetjs_read.js"), sj_dir],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        sys.stderr.write(proc.stderr or proc.stdout)
        raise SystemExit(proc.returncode or 1)
    report = json.load(open(os.path.join(sj_dir, "read.json")))
    return report["version"], {(c["file"], c["cell"]): c for c in report["cases"]}


def cache_matches_got(xml_value, got):
    if isinstance(got, bool) or not isinstance(got, (int, float, str)):
        return False
    if isinstance(got, str):
        return xml_value == got
    try:
        return float(xml_value) == float(got)
    except (TypeError, ValueError):
        return False


def compare_saves(rows):
    """Fill SheetJS write and IronCalc save fields from the written workbooks."""
    by_file = {}
    for row in rows:
        by_file.setdefault(row["file"], []).append((row["sheet"], row["addr"]))
    for file, wanted in by_file.items():
        original = xlsx_xml.cells(os.path.join(FIXTURES, file), wanted)
        sheetjs = xlsx_xml.cells(os.path.join(OUT, "sheetjs", file), wanted)
        ironcalc = xlsx_xml.cells(os.path.join(OUT, "ironcalc", file), wanted)
        for row in rows:
            if row["file"] != file:
                continue
            key = (row["sheet"], row["addr"])
            orig, sj, ic = original[key], sheetjs[key], ironcalc[key]
            values_ok = bool(sj) and bool(orig) and sj["f"] == orig["f"] and sj["v"] == orig["v"]
            meta_ok = bool(sj) and bool(orig) and (sj["fattr"] or {}) == (orig["fattr"] or {}) and sj["cm"] == orig["cm"]
            row["sheetjs_write_ok"] = values_ok
            row["sheetjs_meta_ok"] = meta_ok
            formula_ok = bool(ic) and bool(orig) and ic["f"] == orig["f"]
            if row["error"]:
                cache_ok = False
            elif same(row["got"], row["excel"]):
                cache_ok = bool(ic) and bool(orig) and ic["v"] == orig["v"]
            else:
                # A wrong calculation must be written back, not left as the Excel cache.
                cache_ok = bool(ic) and bool(orig) and cache_matches_got(ic["v"], row["got"]) and ic["v"] != orig["v"]
            row["ironcalc_formula_ok"] = formula_ok
            row["ironcalc_cache_ok"] = cache_ok
            row["ironcalc_cache"] = ic["v"] if ic else None


def classify(rows, baseline):
    for row in rows:
        known = recorded_failure(baseline, row["id"], row["cell"])
        if row["error"]:
            row["status"] = "error"
        elif same(row["got"], row["excel"]):
            row["status"] = "pass"
        elif known is not None and same(row["got"], known):
            row["status"] = "expected-fail"
        else:
            row["status"] = "fail"


# One row of the feature table, then the flag keys that row is built from.
# "bold / fill" is two flags joined with " / " so a change names bold or fill.
FEATURE_ROWS = (
    ("bold / fill", ("bold A1", "fill A1")),
    ("date number format", ("date numfmt A2",)),
    ("column width", ("column width A=30",)),
    ("merged cells", ("merged C1:D1",)),
    ("comment", ("comment A1",)),
    ("data validation", ("data validation B2",)),
    ("freeze panes", ("freeze panes",)),
    ("hyperlink", ("hyperlink B5",)),
    ("defined name", ("defined name MyName",)),
    ("conditional formatting", ("conditional format",)),
    ("chart", ("chart part (xl/charts/)",)),
)
WRITERS = ("IronCalc save", "SheetJS CE write")


def feature_flags(path):
    """Kept/lost signals read from one written RT-features.xlsx.

    openpyxl reads the worksheet, workbook, and styles. The chart check reads
    zip part names. Nothing here is a hand-written kept/lost label.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    a1 = ws["A1"]
    return {
        # xl/styles.xml font <b> on A1's cellXf.
        "bold A1": bool(a1.font and a1.font.b),
        # xl/styles.xml pattern fill <fgColor rgb="..."> on A1's cellXf.
        "fill A1": (a1.fill.fgColor.rgb if a1.fill and a1.fill.fill_type else None) not in (None, "00000000"),
        # xl/styles.xml <numFmt formatCode="yyyy-mm-dd"> on A2.
        "date numfmt A2": ws["A2"].number_format == "yyyy-mm-dd",
        # worksheet <col min="1" max="1" width="..."> within 1.5 of 30.
        "column width A=30": abs((ws.column_dimensions["A"].width or 0) - 30) < 1.5,
        # worksheet <mergeCell ref="C1:D1">.
        "merged C1:D1": "C1:D1" in [str(item) for item in ws.merged_cells.ranges],
        # xl/comments*.xml <comment ref="A1">.
        "comment A1": ws["A1"].comment is not None,
        # worksheet <dataValidations>/<dataValidation> (fixture sqref is B2).
        "data validation B2": len(ws.data_validations.dataValidation) > 0,
        # worksheet <pane state="frozen" topLeftCell="A2">.
        "freeze panes": ws.freeze_panes == "A2",
        # worksheet <hyperlink ref="B5">.
        "hyperlink B5": ws["B5"].hyperlink is not None,
        # workbook.xml <definedName name="MyName">.
        "defined name MyName": "MyName" in wb.defined_names,
        # worksheet <conditionalFormatting> (fixture sqref is A2:A3).
        "conditional format": len(list(ws.conditional_formatting)) > 0,
        # formula string on A4; not a feature-table row.
        "formula A4 '=MyName*2'": ws["A4"].value == "=MyName*2",
        # zip part xl/charts/* (xl/charts/chart1.xml in the fixture).
        "chart part (xl/charts/)": any(name.startswith("xl/charts/") for name in xlsx_xml.parts(path)),
    }


def feature_table(flags):
    rows = []
    for name, keys in FEATURE_ROWS:
        rows.append((name, " / ".join("kept" if flags[key] else "lost" for key in keys)))
    return rows


def _feature_change(name, writer, expected, got):
    if got == "kept" and expected == "lost":
        how = "newly kept"
    elif got == "lost" and expected == "kept":
        how = "newly lost"
    else:
        how = "changed"
    return f"feature {name} {writer}: {how} (baseline {expected}, measured {got})"


def _split_feature_change(name, writer, expected, got):
    if " / " in str(expected) or " / " in str(got):
        labels = name.split(" / ")
        exp_parts = str(expected).split(" / ")
        got_parts = str(got).split(" / ")
        if len(labels) == len(exp_parts) == len(got_parts) and len(labels) > 1:
            return [
                _feature_change(label, writer, exp_part, got_part)
                for label, exp_part, got_part in zip(labels, exp_parts, got_parts)
                if exp_part != got_part
            ]
    return [_feature_change(name, writer, expected, got)]


def feature_mismatches(measured, expected_rows):
    """Name the feature and the library when a kept/lost result changes."""
    mismatches = []
    if len(measured) != len(expected_rows):
        mismatches.append(f"feature rows: measured {len(measured)}, baseline {len(expected_rows)}")
    for got, exp in zip(measured, expected_rows):
        exp_name = exp[0] if exp else None
        if got[0] != exp_name:
            mismatches.append(f"feature row {got[0]!r} != baseline row {exp_name!r}")
            continue
        for index, writer in enumerate(WRITERS, start=1):
            if index >= len(got) or index >= len(exp):
                mismatches.append(f"feature {got[0]} {writer}: missing result")
                continue
            if got[index] != exp[index]:
                mismatches.extend(_split_feature_change(got[0], writer, exp[index], got[index]))
    return mismatches


def ironcalc_wrong_cached_cells(rows):
    """Known misses whose calculated value was written back over the Excel cache."""
    return [
        f"{row['id']} {row['cell']}"
        for row in rows
        if row["status"] == "expected-fail" and row.get("ironcalc_cache_ok")
    ]


def _n_cells(n):
    return f"{n} cell" if n == 1 else f"{n} cells"


def wrong_cache_mismatches(measured, expected):
    if list(measured) == list(expected):
        return []
    measured_text = ", ".join(measured) if measured else "none"
    expected_text = ", ".join(expected) if expected else "none"
    lines = [
        f"IronCalc writes wrong cached values on {_n_cells(len(measured))} ({measured_text}); baseline lists {_n_cells(len(expected))} ({expected_text})"
    ]
    newly = [cell for cell in measured if cell not in expected]
    gone = [cell for cell in expected if cell not in measured]
    if newly:
        lines.append("  newly written wrong: " + ", ".join(newly))
    if gone:
        lines.append("  no longer written wrong: " + ", ".join(gone))
    if not newly and not gone:
        lines.append("  same cells, different order")
    return lines


def counts(rows, baseline):
    weaker = baseline["weaker_fixture"]
    by_id = {}
    for row in rows:
        by_id.setdefault(row["id"], []).append(row)
    return {
        "cells_pass": sum(row["status"] == "pass" for row in rows),
        "cells_pass_excluding_f04": sum(row["status"] == "pass" and row["id"] != weaker for row in rows),
        "cells": len(rows),
        "cells_excluding_f04": sum(row["id"] != weaker for row in rows),
        "fixtures_pass": sum(all(row["status"] == "pass" for row in group) for group in by_id.values()),
        "fixtures": len(by_id),
        "errors": sum(row["status"] == "error" for row in rows),
    }


def failure_line(fixture_id, group):
    got_values = {show(row["got"]) for row in group}
    got_text = next(iter(got_values)) if len(got_values) == 1 else ", ".join(show(row["got"]) for row in group)
    excel_text = ", ".join(show(row["excel"]) for row in group)
    if fixture_id == "F10-sumifs-legacy-whole-column":
        excel_text = " and ".join(show(row["excel"]) for row in group)
        return f"F10 SUMIFS with whole-column criteria returns {got_text} (Excel: {excel_text})"
    if fixture_id == "F07-value-accounting-negatives":
        return f'F07 VALUE of accounting negatives like "(5)" returns {got_text} (Excel: {excel_text})'
    if fixture_id == "F04-text-negative-zero":
        return f'F04 TEXT(-0.4,"0") returns {got_text} vs Excel {excel_text}'
    cells = ", ".join(f"{row['cell']} got {show(row['got'])} excel {show(row['excel'])}" for row in group)
    return f"{fixture_id}: {cells}"


def collect_mismatches(rows, baseline, stats, ic_ver, sj_ver, table, sheetjs_read):
    mismatches = []
    if ic_ver != baseline["ironcalc_version"]:
        mismatches.append(f"IronCalc version {ic_ver} != {baseline['ironcalc_version']}")
    if sj_ver != baseline["sheetjs_version"]:
        mismatches.append(f"SheetJS CE version {sj_ver} != {baseline['sheetjs_version']}")
    for key in ("cells_pass", "cells_pass_excluding_f04", "cells", "cells_excluding_f04", "fixtures_pass", "fixtures", "errors"):
        if stats[key] != baseline[key]:
            mismatches.append(f"{key}: measured {stats[key]}, baseline {baseline[key]}")
    known = {}
    for fixture_id, cells in baseline["ironcalc_failures"].items():
        for cell, got in cells.items():
            known[(fixture_id, cell)] = got
    seen = set()
    for row in rows:
        key = (row["id"], row["cell"])
        if key in known:
            seen.add(key)
            if row["status"] != "expected-fail":
                mismatches.append(
                    f"{row['id']} {row['cell']}: expected IronCalc {show(known[key])}, got {show(row['got']) if row['error'] is None else row['error']}"
                )
        elif row["status"] != "pass":
            mismatches.append(
                f"{row['id']} {row['cell']}: expected Excel {show(row['excel'])}, got {show(row['got']) if row['error'] is None else row['error']}"
            )
        read = sheetjs_read.get((row["file"], row["cell"]))
        if read is None:
            mismatches.append(f"SheetJS read missing {row['file']} {row['cell']}")
        elif not (read["formula_intact"] and read["cached_intact"]):
            mismatches.append(
                f"SheetJS read {row['cell']}: formula {read['read_formula']!r} cached {read['read_cached']!r}"
            )
        if not row.get("sheetjs_write_ok"):
            mismatches.append(f"SheetJS write changed formula or cached value {row['id']} {row['cell']}")
        if not row.get("sheetjs_meta_ok"):
            mismatches.append(f"SheetJS write changed formula attributes or cm metadata {row['id']} {row['cell']}")
        if not row.get("ironcalc_formula_ok"):
            mismatches.append(f"IronCalc save lost formula {row['id']} {row['cell']}")
        if not row.get("ironcalc_cache_ok"):
            mismatches.append(
                f"IronCalc save cache {row['id']} {row['cell']}: wrote {row.get('ironcalc_cache')!r}, calculated {show(row['got']) if row['error'] is None else row['error']}"
            )
    for key in sorted(set(known) - seen):
        mismatches.append(f"baseline failure {key[0]} {key[1]} is not a fixture cell")
    mismatches.extend(feature_mismatches(table, baseline["features"]))
    mismatches.extend(wrong_cache_mismatches(
        ironcalc_wrong_cached_cells(rows),
        baseline["ironcalc_wrong_cached_cells"],
    ))
    return mismatches


def print_js_table(rows, js_pkg, js_ver):
    stats = js_counts(rows)
    print(
        f"IronCalc JS {js_pkg} {js_ver}: "
        f"{stats['cells_match_excel']}/{stats['cells']} cells match Excel, "
        f"{stats['cells_match_python']}/{stats['cells']} match Python, "
        f"{stats['errors']} errors."
    )
    header = ("fixture", "cell", "excel", "python", "js", "js=python", "js=excel")
    body = []
    for row in rows:
        py = row["error"] if row["error"] else show(row["got"])
        js = row["js_error"] if row["js_error"] else show(row["js_got"])
        body.append((
            row["id"],
            row["cell"],
            show(row["excel"]),
            py,
            js,
            "yes" if row["js_matches_python"] else "no",
            "yes" if row["js_matches_excel"] else "no",
        ))
    width = [len(col) for col in header]
    for record in body:
        for i, col in enumerate(record):
            width[i] = max(width[i], len(col))

    def fmt(record):
        return ("  " + "  ".join(col.ljust(width[i]) for i, col in enumerate(record))).rstrip()

    print(fmt(header))
    for record in body:
        print(fmt(record))
    diverged = [row for row in rows if not row["js_matches_python"]]
    if not diverged:
        print("JS vs Python divergences: none")
        return
    print("JS vs Python divergences:")
    for row in diverged:
        py = row["error"] if row["error"] else show(row["got"])
        js = row["js_error"] if row["js_error"] else show(row["js_got"])
        print(f"  {row['id']} {row['cell']}: python {py}, js {js}")


def print_report(ic_ver, sj_ver, rows, stats, table, mismatches, js_pkg, js_ver):
    weaker = "F04"
    print(
        f"IronCalc {ic_ver} (Python bindings) calc parity: "
        f"{stats['cells_pass']}/{stats['cells']} cells pass "
        f"({stats['cells_pass_excluding_f04']}/{stats['cells_excluding_f04']} excluding {weaker}), "
        f"{stats['fixtures_pass']} of {stats['fixtures']} fixtures fully pass, "
        f"{stats['errors']} errors."
    )
    grouped = {}
    for row in rows:
        if row["status"] != "pass":
            grouped.setdefault(row["id"], []).append(row)
    print("Expected failures:")
    ordered = [fid for fid in WEAKER_ORDER if fid in grouped]
    ordered += [fid for fid in grouped if fid not in ordered]
    if not ordered:
        print("  (none)")
    for fid in ordered:
        print(f"  {failure_line(fid, grouped[fid])}")
    read_ok = sum(bool(row.get("_read_ok")) for row in rows)
    write_ok = sum(bool(row.get("sheetjs_write_ok")) for row in rows)
    formula_ok = sum(bool(row.get("ironcalc_formula_ok")) for row in rows)
    print(f"SheetJS CE {sj_ver} read: {read_ok}/{stats['cells']} formula strings and cached values preserved.")
    print(f"SheetJS CE {sj_ver} write: {write_ok}/{stats['cells']} formula strings and cached values preserved.")
    wrong_cells = ironcalc_wrong_cached_cells(rows)
    print(
        f"IronCalc save: {formula_ok}/{stats['cells']} formula strings kept; "
        f"writes wrong cached values on {len(wrong_cells)} cells."
    )
    for cell in wrong_cells:
        print(f"  {cell}")
    print("Feature round-trip (RT-features.xlsx):")
    header = ("feature", "IronCalc save", "SheetJS CE write")
    width = [len(header[i]) for i in range(3)]
    body = [tuple(row) for row in table]
    for row in body:
        for i, col in enumerate(row):
            width[i] = max(width[i], len(col))
    def fmt(row):
        return ("  " + "  ".join(col.ljust(width[i]) for i, col in enumerate(row))).rstrip()
    print(fmt(header))
    for row in body:
        print(fmt(row))
    print()
    print("cells:")
    for row in rows:
        sj = "ok" if row.get("_read_ok") else "LOST"
        got = row["error"] if row["error"] else show(row["got"])
        print(
            f"  {row['status']:<14} {row['id']} {row['cell']}  "
            f"{human(row['formula'])}  excel={show(row['excel'])} ironcalc={got} sheetjs-read={sj}"
        )
    print()
    print_js_table(rows, js_pkg, js_ver)
    print()
    if mismatches:
        print("baseline: DIFFERS")
        for item in mismatches:
            print(f"  {item}")
    else:
        print("baseline: match")


def main():
    os.makedirs(OUT, exist_ok=True)
    fixtures = load_fixtures()
    baseline = load_baseline()
    ic_ver, rows = run_ironcalc(fixtures)
    js_pkg, js_ver, js_cases = run_ironcalc_js()
    attach_js(rows, js_cases)
    sj_ver, sheetjs_read = run_sheetjs()
    for row in rows:
        read = sheetjs_read.get((row["file"], row["cell"]))
        row["_read_ok"] = bool(read and read["formula_intact"] and read["cached_intact"])
    compare_saves(rows)
    classify(rows, baseline)
    stats = counts(rows, baseline)
    ic_flags = feature_flags(os.path.join(OUT, "ironcalc", "RT-features.xlsx"))
    sj_flags = feature_flags(os.path.join(OUT, "sheetjs", "RT-features.xlsx"))
    orig_flags = feature_flags(os.path.join(FIXTURES, "RT-features.xlsx"))
    ic_table = feature_table(ic_flags)
    sj_table = feature_table(sj_flags)
    table = [[name, ic, sj] for (name, ic), (_, sj) in zip(ic_table, sj_table)]
    mismatches = collect_mismatches(rows, baseline, stats, ic_ver, sj_ver, table, sheetjs_read)
    mismatches.extend(collect_js_mismatches(rows, baseline, js_pkg, js_ver))
    for name, value in feature_table(orig_flags):
        if value not in ("kept", "kept / kept"):
            mismatches.append(f"original RT-features.xlsx missing {name} ({value})")
    print_report(ic_ver, sj_ver, rows, stats, table, mismatches, js_pkg, js_ver)
    return 1 if mismatches else 0


if __name__ == "__main__":
    sys.exit(main())
