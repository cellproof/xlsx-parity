"""Score LibreOffice Calc on the formula fixtures.

``soffice --headless --convert-to xlsx`` keeps the cached values stored in the
file unless the profile sets OOXMLRecalcMode to 0 (always recalculate on
load). The default is ask, and a headless convert treats ask as never, so a
plain convert echoes the Excel cache. This run deletes formula caches in a
temporary copy, converts with that setting, and reads the new ``<v>`` values.

F04 is converted a second time with its original cache and no recalc setting.
That control is the echo. The cleared-cache convert is the calculation.
"""
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

import xlsx_xml

NSM = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NSM)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES = os.path.join(ROOT, "fixtures")
PROOF_FILE = "F04-text-negative-zero.xlsx"
PROOF_CELL = ("Sheet1", "A1")

# 0 = always recalculate OOXML formula caches on load.
RECALC_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def same_value(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a is b
    if isinstance(a, str) or isinstance(b, str):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b
    return a == b


def from_cache(cell):
    """Python value for one cached ``<c>``. Numbers stay numbers; bools stay bools."""
    if not cell or cell.get("v") is None:
        return None
    kind = cell.get("t")
    raw = cell["v"]
    if kind == "b":
        return raw == "1"
    if kind == "e":
        return raw
    if kind in ("str", "s", "inlineStr"):
        return raw
    number = float(raw)
    if number.is_integer() and "." not in raw and "e" not in raw.lower():
        return int(number)
    return number


def product_version():
    if shutil.which("soffice") is None:
        raise SystemExit(
            "LibreOffice Calc is required (soffice is not on PATH). "
            "Install libreoffice-calc 4:26.2.5.2-0ubuntu0.26.04.1~bpo24.04.1; see README."
        )
    proc = subprocess.run(["soffice", "--version"], capture_output=True, text=True, timeout=60)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "soffice --version failed")
    line = proc.stdout.strip().splitlines()
    if not line:
        raise RuntimeError("soffice --version printed nothing")
    return line[0]


def package_version():
    proc = subprocess.run(
        ["dpkg-query", "-W", "-f=${Version}", "libreoffice-calc"],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        return None
    return proc.stdout.strip() or None


def lo_env(home):
    env = os.environ.copy()
    env["HOME"] = home
    # Avoid a GTK/display plugin on a headless runner.
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


def make_profile(path, recalc):
    user = os.path.join(path, "user")
    os.makedirs(user, exist_ok=True)
    if recalc:
        with open(os.path.join(user, "registrymodifications.xcu"), "w", encoding="utf-8") as handle:
            handle.write(RECALC_XCU)


def run_soffice(args, env, timeout):
    proc = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        start_new_session=True,
        env=env,
    )
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        os.killpg(proc.pid, signal.SIGKILL)
        proc.communicate()
        raise RuntimeError(f"soffice timed out after {timeout}s")
    if proc.returncode != 0:
        detail = (stderr or stdout or "").strip()
        raise RuntimeError(f"soffice exit {proc.returncode}: {detail}")
    if "Error:" in stdout or "Error:" in stderr:
        detail = (stderr or stdout).strip()
        raise RuntimeError(f"soffice reported an error: {detail}")
    return stdout


def convert_cmd(profile, outdir, files):
    return [
        "soffice",
        f"-env:UserInstallation=file://{os.path.abspath(profile)}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "xlsx",
        "--outdir",
        outdir,
        *files,
    ]


def blank_formula_caches(src, dest):
    """Drop cached formula results so a convert cannot echo them. Return True if none remain."""
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dest, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml"):
                root = ET.fromstring(data)
                for cell in root.iter(f"{{{NSM}}}c"):
                    if cell.find(f"{{{NSM}}}f") is None:
                        continue
                    for tag in ("v", "is"):
                        child = cell.find(f"{{{NSM}}}{tag}")
                        if child is not None:
                            cell.remove(child)
                    cell.attrib.pop("t", None)
                data = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
            zout.writestr(info.filename, data)
    with zipfile.ZipFile(dest) as check:
        for name in check.namelist():
            if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
                continue
            root = ET.fromstring(check.read(name))
            for cell in root.iter(f"{{{NSM}}}c"):
                if cell.find(f"{{{NSM}}}f") is not None and (
                    cell.find(f"{{{NSM}}}v") is not None or cell.find(f"{{{NSM}}}is") is not None
                ):
                    return False
    return True


def _read_one(path, sheet, addr):
    found = xlsx_xml.cells(path, [(sheet, addr)])
    return from_cache(found.get((sheet, addr)))


def run(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    product = product_version()
    deb = package_version()
    fixtures = json.load(open(os.path.join(FIXTURES, "fixtures.json"), encoding="utf-8"))
    started = time.time()
    tmp = tempfile.mkdtemp(prefix="lo-calc-")
    try:
        control_profile = os.path.join(tmp, "control-profile")
        control_out = os.path.join(tmp, "control-out")
        os.makedirs(control_out)
        make_profile(control_profile, False)
        proof_src = os.path.join(FIXTURES, PROOF_FILE)
        run_soffice(
            convert_cmd(control_profile, control_out, [proof_src]),
            lo_env(os.path.join(tmp, "control-home")),
            60,
        )

        src_dir = os.path.join(tmp, "src")
        os.makedirs(src_dir)
        caches_cleared = True
        for fx in fixtures:
            dest = os.path.join(src_dir, fx["file"])
            if not blank_formula_caches(os.path.join(FIXTURES, fx["file"]), dest):
                caches_cleared = False
        recalc_profile = os.path.join(tmp, "recalc-profile")
        calc_out = os.path.join(out_dir, "xlsx")
        os.makedirs(calc_out, exist_ok=True)
        for old in os.listdir(calc_out):
            if old.endswith(".xlsx"):
                os.remove(os.path.join(calc_out, old))
        make_profile(recalc_profile, True)
        sources = [os.path.join(src_dir, fx["file"]) for fx in fixtures]
        run_soffice(
            convert_cmd(recalc_profile, calc_out, sources),
            lo_env(os.path.join(tmp, "home")),
            180,
        )

        cases = []
        differ = 0
        for fx in fixtures:
            wanted = [(case[0], case[1]) for case in fx["cases"]]
            written = os.path.join(calc_out, fx["file"])
            if not os.path.exists(written):
                for sheet, addr, _formula, _expected, _source in fx["cases"]:
                    cases.append({
                        "file": fx["file"],
                        "cell": f"{sheet}!{addr}",
                        "got": None,
                        "error": f"LibreOffice did not write {fx['file']}",
                    })
                continue
            got_cells = xlsx_xml.cells(written, wanted)
            orig_cells = xlsx_xml.cells(os.path.join(FIXTURES, fx["file"]), wanted)
            for sheet, addr, _formula, _expected, _source in fx["cases"]:
                ref = f"{sheet}!{addr}"
                cell = got_cells.get((sheet, addr))
                if cell is None or cell.get("v") is None:
                    cases.append({
                        "file": fx["file"],
                        "cell": ref,
                        "got": None,
                        "error": "no cached value after recalc",
                    })
                    continue
                got = from_cache(cell)
                cases.append({"file": fx["file"], "cell": ref, "got": got, "error": None})
                if not same_value(got, from_cache(orig_cells.get((sheet, addr)))):
                    differ += 1
        sheet, addr = PROOF_CELL
        proof_ref = f"{sheet}!{addr}"
        scored = next(case for case in cases if case["file"] == PROOF_FILE and case["cell"] == proof_ref)
        report = {
            "product": product,
            "package": "libreoffice-calc",
            "package_version": deb,
            "seconds": round(time.time() - started, 2),
            "caches_cleared": caches_cleared,
            "cells_differ_from_cache": differ,
            "recalc_proof": {
                "file": PROOF_FILE,
                "cell": proof_ref,
                "stored_cache": _read_one(proof_src, sheet, addr),
                "without_recalc": _read_one(os.path.join(control_out, PROOF_FILE), sheet, addr),
                "with_recalc": scored.get("error") or scored.get("got"),
            },
            "cases": cases,
        }
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    with open(os.path.join(out_dir, "read.json"), "w", encoding="utf-8") as handle:
        json.dump(report, handle)
    return report


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "out", "libreoffice")
    result = run(out)
    proof = result["recalc_proof"]
    print(
        f"{result['product']} {result['package']} {result['package_version']} "
        f"{result['seconds']}s differ={result['cells_differ_from_cache']} "
        f"proof {proof['cell']} cache={proof['stored_cache']!r} "
        f"plain={proof['without_recalc']!r} recalc={proof['with_recalc']!r}"
    )
