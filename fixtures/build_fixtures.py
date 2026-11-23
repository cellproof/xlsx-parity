"""Rebuild the 10 formula fixtures and fixtures.json.

make test uses the committed workbooks. Run this only to regenerate them.
Every expected value is copied from the cited Univer issue or Microsoft Support
page. Each formula cell stores that Excel value as its cached value, so a
reader that does not calculate can be checked for cache survival.

Formula strings are OOXML storage form (_xlfn. / _xlpm. where Excel uses them).
"""
import json
import os

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
UNIVER = "https://github.com/dream-num/univer/issues/"
MS = {
    "ISO.CEILING": "https://support.microsoft.com/en-us/excel/functions/iso-ceiling-function",
    "EDATE": "https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5",
    "EOMONTH": "https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628",
    "SEARCH": "https://support.microsoft.com/en-us/office/search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495",
    "CHAR": "https://support.microsoft.com/en-us/office/char-function-bbd249c8-b36e-4a91-8017-1c133f9b837a",
    "AMORDEGRC": "https://support.microsoft.com/en-us/office/amordegrc-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51",
    "IMPLICIT": "https://support.microsoft.com/en-us/office/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34",
}

# case = (sheet, cell, stored_formula, expected, expected_source_note)
FIXTURES = [
    dict(id="F01-let-lambda-call", issue=8332,
         sources=[UNIVER + "8332"],
         note="Issue table; author says the Excel column was checked in Excel for the web.",
         inputs={}, cases=[
            ("Sheet1", "A1", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.p,_xlpm.p+1),_xlpm.f(1))", 2, "issue #8332 table"),
            ("Sheet1", "A2", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.p,_xlpm.q,_xlpm.p+_xlpm.q),_xlpm.f(1,2))", 3, "issue #8332 table"),
         ]),
    dict(id="F02-iso-ceiling", issue=8268,
         sources=[UNIVER + "8268", MS["ISO.CEILING"]],
         note="Expected values stated in the issue, citing Microsoft's ISO.CEILING page.",
         inputs={}, cases=[
            ("Sheet1", "A1", "_xlfn.ISO.CEILING(4.3)", 5, "issue #8268 / MS ISO.CEILING"),
            ("Sheet1", "A2", "_xlfn.ISO.CEILING(-4.3)", -4, "issue #8268 / MS ISO.CEILING"),
            ("Sheet1", "A3", "_xlfn.ISO.CEILING(4.3,-2)", 6, "issue #8268 / MS ISO.CEILING"),
         ]),
    dict(id="F03-edate-eomonth-negative-fraction", issue=8269,
         sources=[UNIVER + "8269", MS["EDATE"], MS["EOMONTH"]],
         note="Serials 40527 / 40543 stated in the issue; MS docs say non-integer months are truncated.",
         inputs={}, date_format=True, cases=[
            ("Sheet1", "A1", "EDATE(DATE(2011,1,15),-1.5)", 40527, "issue #8269 / MS EDATE"),
            ("Sheet1", "A2", "EOMONTH(DATE(2011,1,15),-1.5)", 40543, "issue #8269 / MS EOMONTH"),
         ]),
    dict(id="F04-text-negative-zero", issue=8270,
         sources=[UNIVER + "8270"],
         note="Expected '-0' is the issue author's statement of Excel behavior; the author says they could not run it in Excel. Sourced from issue text, but weaker than the others.",
         inputs={}, cases=[
            ("Sheet1", "A1", 'TEXT(-0.4,"0")', "-0", "issue #8270 (author unverified in Excel)"),
         ]),
    dict(id="F05-search-wildcards-and-round-digits", issue=8198,
         sources=[UNIVER + "8198", MS["SEARCH"]],
         note="SEARCH: issue says all calls return 2 per MS SEARCH docs. ROUND*: issue says Excel returns the number unchanged or 0.",
         inputs={}, cases=[
            ("Sheet1", "A1", 'SEARCH("b*","abc")', 2, "issue #8198 / MS SEARCH"),
            ("Sheet1", "A2", 'SEARCH("?c","abc")', 2, "issue #8198 / MS SEARCH"),
            ("Sheet1", "A3", 'SEARCH("a*c","xabc")', 2, "issue #8198 / MS SEARCH"),
            ("Sheet1", "A4", 'SEARCH("~*","a*b")', 2, "issue #8198 / MS SEARCH"),
            ("Sheet1", "B1", "ROUNDUP(1.23,16)", 1.23, "issue #8198 text"),
            ("Sheet1", "B2", "ROUND(1.5,20)", 1.5, "issue #8198 text"),
            ("Sheet1", "B3", "ROUNDDOWN(1.5,309)", 1.5, "issue #8198 text"),
            ("Sheet1", "B4", "ROUND(2.5,-400)", 0, "issue #8198 text"),
         ]),
    dict(id="F06-char-range", issue=8211,
         sources=[UNIVER + "8211", MS["CHAR"]],
         note="MS CHAR takes 'a number between 1 and 255'; issue states CHAR(256) and CHAR(65601) return #VALUE!. 128-255 is platform-dependent, so not tested.",
         inputs={}, cases=[
            ("Sheet1", "A1", "CHAR(256)", "#VALUE!", "issue #8211 / MS CHAR"),
            ("Sheet1", "A2", "CHAR(65601)", "#VALUE!", "issue #8211 / MS CHAR"),
            ("Sheet1", "A3", "CHAR(65)", "A", "MS CHAR (65 -> A is in the 1-255 ASCII range)"),
         ]),
    dict(id="F07-value-accounting-negatives", issue=8231,
         sources=[UNIVER + "8231"],
         note="Expected values stated in the issue text as Excel behavior.",
         inputs={}, cases=[
            ("Sheet1", "A1", 'VALUE("(5)")', -5, "issue #8231 text"),
            ("Sheet1", "A2", 'VALUE("($1,000)")', -1000, "issue #8231 text"),
            ("Sheet1", "A3", '"(5)"+1', -4, "issue #8231 text"),
         ]),
    dict(id="F08-amordegrc", issue=8267,
         sources=[UNIVER + "8267", MS["AMORDEGRC"]],
         note="#8267 lists AMORDEGRC as evaluating to #NAME? in Univer. Expected 776 is the worked example on Microsoft's AMORDEGRC page.",
         inputs={("Sheet1", "A2"): 2400, ("Sheet1", "A3"): 39679, ("Sheet1", "A4"): 39813,
                 ("Sheet1", "A5"): 300, ("Sheet1", "A6"): 1, ("Sheet1", "A7"): 0.15, ("Sheet1", "A8"): 1},
         cases=[
            ("Sheet1", "A10", "AMORDEGRC(A2,A3,A4,A5,A6,A7,A8)", 776, "MS AMORDEGRC example"),
         ]),
    dict(id="F09-ast-cache-collision-cells", issue=7986,
         sources=[UNIVER + "7986"],
         note="Exact cells from the issue; issue states Q16084 should be 328105 (Univer showed 26614).",
         inputs={("Sheet1", "A1"): 10},
         cases=[
            ("Sheet1", "P1169", "A1*326+23354", 26614, "issue #7986 (=10*326+23354)"),
            ("Sheet1", "Q16084", "A1*645+321655", 328105, "issue #7986 text"),
         ]),
    dict(id="F10-sumifs-legacy-whole-column", issue=8154,
         sources=[UNIVER + "8154", MS["IMPLICIT"]],
         note=("Issue's own data: Data!A = k{i%300}, Data!B = i+1 for 2,500 rows; Summary!A = k0..k299; "
               "Summary!B = SUMIFS(Data!B:B,Data!A:A,$A:$A) stored WITHOUT @ (legacy, non-array), which is how "
               "pre-dynamic-array Excel files store it. Per the issue and MS 'Implicit intersection operator', Excel "
               "shows it as @$A:$A and returns the same-row total. Issue states k0 total 10809, k299 total 10800. "
               "Cached values for other rows are computed from the same data definition and are not asserted."),
         inputs="sumifs", cases=[
            ("Summary", "B2", "SUMIFS(Data!B:B,Data!A:A,$A:$A)", 10809, "issue #8154 text (k0 total)"),
            ("Summary", "B301", "SUMIFS(Data!B:B,Data!A:A,$A:$A)", 10800, "issue #8154 text (k299 total)"),
         ]),
]


def build(fx, path):
    wb = xlsxwriter.Workbook(path)
    datefmt = wb.add_format({"num_format": "yyyy-mm-dd"})
    sheets = {}

    def sh(name):
        if name not in sheets:
            sheets[name] = wb.add_worksheet(name)
        return sheets[name]

    if fx["inputs"] == "sumifs":
        s = sh("Summary")
        d = sh("Data")
        d.write("A1", "key")
        d.write("B1", "amount")
        totals = {}
        for i in range(2500):
            d.write_string(i + 1, 0, f"k{i % 300}")
            d.write_number(i + 1, 1, i + 1)
            totals[i % 300] = totals.get(i % 300, 0) + i + 1
        s.write("A1", "key")
        s.write("B1", "total")
        for i in range(300):
            s.write_string(i + 1, 0, f"k{i}")
            s.write_formula(i + 1, 1, "=SUMIFS(Data!B:B,Data!A:A,$A:$A)", None, totals[i])
    else:
        for (sn, cell), val in fx["inputs"].items():
            sh(sn).write(cell, val)
        for sn, cell, formula, expected, _src in fx["cases"]:
            fmt = datefmt if fx.get("date_format") else None
            sh(sn).write_formula(cell, "=" + formula, fmt, expected)
    wb.close()


if __name__ == "__main__":
    manifest = []
    for fx in FIXTURES:
        name = fx["id"] + ".xlsx"
        build(fx, os.path.join(HERE, name))
        manifest.append({k: v for k, v in fx.items() if k != "inputs"} | {"file": name})
    with open(os.path.join(HERE, "fixtures.json"), "w") as fh:
        json.dump(manifest, fh, indent=1)
        fh.write("\n")
    print("wrote", len(FIXTURES), "fixtures,", sum(len(f["cases"]) for f in FIXTURES), "cases")
