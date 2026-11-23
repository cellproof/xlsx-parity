"""Test 1 baseline: IronCalc calculation, SheetJS CE read/write, feature round-trip.

Excel-expected values and their citations live in fixtures/fixtures.json.
Known IronCalc misses are recorded in suite/baseline.json. The process exits 0
when every measured result matches that record, and nonzero when a result
changes — a regression or an unexpected fix.
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


def feature_flags(path):
    flags = {}
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    a1 = ws["A1"]
    flags["bold A1"] = bool(a1.font and a1.font.b)
    flags["fill A1"] = (a1.fill.fgColor.rgb if a1.fill and a1.fill.fill_type else None) not in (None, "00000000")
    flags["date numfmt A2"] = ws["A2"].number_format == "yyyy-mm-dd"
    flags["column width A=30"] = abs((ws.column_dimensions["A"].width or 0) - 30) < 1.5
    flags["merged C1:D1"] = "C1:D1" in [str(x) for x in ws.merged_cells.ranges]
    flags["comment A1"] = ws["A1"].comment is not None
    flags["data validation B2"] = len(ws.data_validations.dataValidation) > 0
    flags["freeze panes"] = ws.freeze_panes == "A2"
    flags["hyperlink B5"] = ws["B5"].hyperlink is not None
    flags["defined name MyName"] = "MyName" in wb.defined_names
    flags["conditional format"] = len(list(ws.conditional_formatting)) > 0
    flags["formula A4 '=MyName*2'"] = ws["A4"].value == "=MyName*2"
    flags["chart part (xl/charts/)"] = any(name.startswith("xl/charts/") for name in xlsx_xml.parts(path))
    return flags


def feature_table(flags):
    kept = lambda name: "kept" if flags[name] else "lost"
    return [
        ("bold / fill", f"{kept('bold A1')} / {kept('fill A1')}"),
        ("date number format", kept("date numfmt A2")),
        ("column width", kept("column width A=30")),
        ("merged cells", kept("merged C1:D1")),
        ("comment", kept("comment A1")),
        ("data validation", kept("data validation B2")),
        ("freeze panes", kept("freeze panes")),
        ("hyperlink", kept("hyperlink B5")),
        ("defined name", kept("defined name MyName")),
        ("conditional formatting", kept("conditional format")),
        ("chart", kept("chart part (xl/charts/)")),
    ]


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
    if table != baseline["features"]:
        mismatches.append("feature table differs from baseline")
        for got, exp in zip(table, baseline["features"]):
            if list(got) != exp:
                mismatches.append(f"  measured {got} baseline {exp}")
        if len(table) != len(baseline["features"]):
            mismatches.append(f"  measured {len(table)} feature rows, baseline {len(baseline['features'])}")
    return mismatches


def print_report(ic_ver, sj_ver, rows, stats, table, mismatches):
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
    overwritten = sum(row["status"] == "expected-fail" and row.get("ironcalc_cache_ok") for row in rows)
    print(
        f"IronCalc save: {formula_ok}/{stats['cells']} formula strings kept; "
        f"writes wrong cached values on {overwritten} cells."
    )
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
    for name, value in feature_table(orig_flags):
        if value not in ("kept", "kept / kept"):
            mismatches.append(f"original RT-features.xlsx missing {name} ({value})")
    print_report(ic_ver, sj_ver, rows, stats, table, mismatches)
    return 1 if mismatches else 0


if __name__ == "__main__":
    sys.exit(main())
