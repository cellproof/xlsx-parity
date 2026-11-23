"""Rebuild the formula fixtures and fixtures.json.

make test uses the committed workbooks. Run this only to regenerate them.
Every expected value is copied from a cited source: a Univer issue, an
IronCalc issue, HyperFormula's published Excel column, or a Microsoft
Support page that states the result for that input. Each formula cell
stores that Excel value as its cached value, so a reader that does not
calculate can be checked for cache survival.

Formula strings are OOXML storage form (_xlfn. / _xlpm. / _xlws. where
Excel uses them).
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
    "INT": "https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef",
    "ISEVEN": "https://support.microsoft.com/en-us/office/iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356",
}
HF = "https://hyperformula.handsontable.com/docs/guide/list-of-differences"
IC = "https://github.com/ironcalc/IronCalc/issues/"

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
    # Milestone 4. Excel values below are copied from HyperFormula's differences
    # table or from an IronCalc issue that shows Excel's output. None are inferred.
    dict(id="F11-hf-negative-rounding",
         sources=[HF, MS["INT"], MS["ISEVEN"]],
         note=("HyperFormula differences table, negative-rounding rows. "
               "INT(-8.9)=-9 and ISEVEN(2.5)=TRUE are also the worked examples on the Microsoft pages."),
         inputs={}, cases=[
            ("Sheet1", "A1", "INT(-8.9)", -9, "HF differences / MS INT"),
            ("Sheet1", "A2", "ISEVEN(2.5)", True, "HF differences / MS ISEVEN"),
            ("Sheet1", "A3", "ISODD(3.5)", True, "HF differences table (Excel TRUE)"),
            ("Sheet1", "A4", "_xlfn.CEILING.MATH(-4.3,2,2)", -6, "HF differences table"),
            ("Sheet1", "A5", "_xlfn.FLOOR.MATH(-4.7,2,2)", -4, "HF differences table"),
         ]),
    dict(id="F12-hf-npv-sequence-errors",
         sources=[HF],
         note=("Excel column of HyperFormula's differences table. Rounded 4-decimal rows are not included. "
               "DAYS is stored without _xlfn.: IronCalc 0.9.1 rewrites the prefixed form to DAYS on save."),
         inputs={("Sheet1", "B1"): True}, cases=[
            ("Sheet1", "A1", "NPV(1,TRUE(),1)", 0.75, "HF differences table"),
            ("Sheet1", "A2", "NPV(1,B1)", 0, "HF differences table (B1 = TRUE)"),
            ("Sheet1", "A3", "COUNTA(1,)", 2, "HF differences table"),
            ("Sheet1", "A4", "_xlfn.T.INV(0,1)", "#DIV/0!", "HF differences table"),
            ("Sheet1", "A5", "_xlfn.SEQUENCE(0)", "#CALC!", "HF differences table"),
            ("Sheet1", "A6", "_xlfn.WEIBULL.DIST(0,1,1,FALSE())", 0, "HF differences table"),
            ("Sheet1", "A7", 'ADDRESS(1,1,4,TRUE(),"")', "!A1", "HF differences table"),
            ("Sheet1", "A8", "DAYS(-1,0)", "#NUM!", "HF differences table"),
         ]),
    dict(id="F13-hf-if-false-cycle",
         sources=[HF],
         note="A1 refers to itself only on the branch IF does not take. HF says Excel returns 0, same as Google Sheets.",
         inputs={}, cases=[
            ("Sheet1", "A1", "IF(FALSE(),A1,0)", 0, "HF differences table (dependency collection)"),
         ]),
    dict(id="F14-large-k", issue=1750,
         sources=[IC + "1750"],
         note="B1:B4 is 1, 3, 5, 8. Author ran these in Excel for Microsoft 365 (16.0.20527.20150).",
         inputs={("Sheet1", "B1"): 1, ("Sheet1", "B2"): 3, ("Sheet1", "B3"): 5, ("Sheet1", "B4"): 8},
         cases=[
            ("Sheet1", "A1", "LARGE(B1:B4,1.9)", 5, "issue #1750 table"),
            ("Sheet1", "A2", "LARGE(B1:B4,4.5)", "#NUM!", "issue #1750 table"),
            ("Sheet1", "A3", "LARGE(B1:B4,TRUE)", 8, "issue #1750 table"),
            ("Sheet1", "A4", "SMALL(B1:B4,TRUE)", 1, "issue #1750 table"),
         ]),
    dict(id="F15-vlookup-index-below-one", issue=1751,
         sources=[IC + "1751"],
         note="A1:C2 is 1, one, 10 / 2, two, 20. Author ran these in Excel for Microsoft 365 (16.0.20527.20150).",
         inputs={("Sheet1", "A1"): 1, ("Sheet1", "B1"): "one", ("Sheet1", "C1"): 10,
                 ("Sheet1", "A2"): 2, ("Sheet1", "B2"): "two", ("Sheet1", "C2"): 20},
         cases=[
            ("Sheet1", "E1", "VLOOKUP(2,A1:C2,0,FALSE)", "#VALUE!", "issue #1751 table"),
            ("Sheet1", "E2", "HLOOKUP(1,A1:C2,0.5,FALSE)", "#VALUE!", "issue #1751 table"),
         ]),
    dict(id="F16-text-general", issue=1752,
         sources=[IC + "1752"],
         note="Author ran these in Excel for Microsoft 365 (16.0.20527.20150). TEXT(0.000000001,\"General\") is omitted: IronCalc rewrites that literal on save.",
         inputs={}, cases=[
            ("Sheet1", "A1", 'TEXT(1/33600,"General")', "2.97619E-05", "issue #1752 table"),
            ("Sheet1", "A2", 'TEXT(12345678.9012345,"General")', "12345678.9", "issue #1752 table"),
            ("Sheet1", "A3", 'TEXT(100000000000/3,"General")', "33333333333", "issue #1752 table"),
         ]),
    dict(id="F17-explicit-argument-coercion", issue=1605,
         sources=[IC + "1605"],
         note="Excel values are the cached results the reporter copied from an Excel-saved POI workbook (EverythingTests).",
         inputs={}, cases=[
            ("Sheet1", "A1", 'LEFT("a string","6")', "a stri", "issue #1605 table"),
            ("Sheet1", "A2", 'SUMSQ("2","3",4)', 29, "issue #1605 table"),
            ("Sheet1", "A3", 'STDEV("1","2")', 0.7071067811865476, "issue #1605 table"),
            ("Sheet1", "A4", "VAR(TRUE,FALSE,TRUE)", 0.33333333333333337, "issue #1605 table"),
         ]),
    dict(id="F18-xlookup-return-and-empty",
         sources=[IC + "1789", IC + "1718"],
         note="A1:C3 and D1 are issue #1789. E1:F4 and G1 are issue #1718. Both tables state Excel's cell value.",
         inputs={
             ("Sheet1", "A1"): "b", ("Sheet1", "A2"): "a", ("Sheet1", "A3"): "b",
             ("Sheet1", "B1"): "x", ("Sheet1", "C1"): "y",
             ("Sheet1", "B2"): "p", ("Sheet1", "C2"): "q",
             ("Sheet1", "B3"): "r", ("Sheet1", "C3"): "s",
             ("Sheet1", "E1"): 5, ("Sheet1", "E2"): 3, ("Sheet1", "E3"): 8, ("Sheet1", "E4"): 1,
             ("Sheet1", "F1"): "b", ("Sheet1", "F2"): "d", ("Sheet1", "F3"): "a", ("Sheet1", "F4"): "c",
         }, cases=[
            ("Sheet1", "D1", '_xlfn.XLOOKUP("b",A1:A3,B1:C3)', "x", "issue #1789 table (first spill cell)"),
            ("Sheet1", "G1", "_xlfn.XLOOKUP(8,E1:E4,F1:F4,,,)", "a", "issue #1718 table"),
         ]),
    dict(id="F19-sort-empty-argument", issue=1597,
         sources=[IC + "1597"],
         note="A1:A4 is 5, 3, 8, 1. The issue table states Excel's results for the empty optional arguments.",
         inputs={("Sheet1", "A1"): 5, ("Sheet1", "A2"): 3, ("Sheet1", "A3"): 8, ("Sheet1", "A4"): 1},
         cases=[
            ("Sheet1", "B1", "INDEX(_xlfn._xlws.SORT(A1:A4,,-1),1)", 8, "issue #1597 table"),
            ("Sheet1", "B2", "_xlfn.XMATCH(8,A1:A4,,)", 3, "issue #1597 table"),
         ]),
    dict(id="F20-date-serial-zero", issue=1569,
         sources=[IC + "1569"],
         note="Author's table from Excel for Mac 16.114. Serial 0 is accepted; DAY is 0 and WEEKDAY is 7.",
         inputs={}, cases=[
            ("Sheet1", "A1", "YEAR(0)", 1900, "issue #1569 table"),
            ("Sheet1", "A2", "DAY(0)", 0, "issue #1569 table"),
            ("Sheet1", "A3", "WEEKDAY(0)", 7, "issue #1569 table"),
            ("Sheet1", "A4", "YEAR(B1)", 1900, "issue #1569 table (B1 blank)"),
         ]),
    dict(id="F21-date-year-offset", issue=1364,
         sources=[IC + "1364"],
         note="Years 0 through 1899 are added to 1900. The issue table states the serials.",
         inputs={}, cases=[
            ("Sheet1", "A1", "DATE(1,1,1)", 367, "issue #1364 table"),
            ("Sheet1", "A2", "DATE(100,1,1)", 36526, "issue #1364 table"),
         ]),
    dict(id="F22-and-empty-text", issue=1609,
         sources=[IC + "1609"],
         note="A1 and B1 are empty text. The issue confirms this shape as #VALUE! from an Excel-saved workbook. The 1/0 rows in that issue were not run in Excel and are not here.",
         inputs={}, helpers=[
            ("Sheet1", "A1", '""', ""),
            ("Sheet1", "B1", '""', ""),
         ], cases=[
            ("Sheet1", "C1", 'AND(A1<>"",A1+B1<1)', "#VALUE!", "issue #1609 table"),
         ]),
    dict(id="F23-sum-if-range", issue=1749,
         sources=[IC + "1749"],
         note="B1:B4 is 1, 3, 5, 8 and C1:C4 is 10, 20, 30, 40. Formulas sit outside those rows. Author ran Excel for Microsoft 365 (16.0.20527.20150).",
         inputs={
             ("Sheet1", "B1"): 1, ("Sheet1", "B2"): 3, ("Sheet1", "B3"): 5, ("Sheet1", "B4"): 8,
             ("Sheet1", "C1"): 10, ("Sheet1", "C2"): 20, ("Sheet1", "C3"): 30, ("Sheet1", "C4"): 40,
         }, cases=[
            ("Sheet1", "E6", "SUM(IF(TRUE,B1:B4))", 17, "issue #1749 table"),
            ("Sheet1", "E7", "SUM(IF(B1>5,B1:B4,C1:C4))", 100, "issue #1749 table"),
         ]),
    dict(id="F24-subtotal-edges", issue=1739,
         sources=[IC + "1739"],
         note=("Author ran both tables in Excel for Microsoft 365 (16.0.20527.20150). "
               "Row 2 is hidden. H1:H3 is 2, 8, 4. D6 holds a SUBTOTAL that is not the whole formula."),
         hidden_rows=[("Sheet1", 1)],
         inputs={("Sheet1", "D3"): 1, ("Sheet1", "D4"): 2, ("Sheet1", "D5"): 3,
                 ("Sheet1", "H1"): 2, ("Sheet1", "H2"): 8, ("Sheet1", "H3"): 4},
         helpers=[
            ("Sheet1", "D6", "SUBTOTAL(9,D3:D5)+1", 7),
         ], cases=[
            ("Sheet1", "A1", "SUBTOTAL(6,B1:B2)", 0, "issue #1739 table (product of no numbers)"),
            ("Sheet1", "A8", "SUBTOTAL(110,H1:H3)", 2, "issue #1739 table (VAR.S, row 2 hidden)"),
            ("Sheet1", "A9", "SUBTOTAL(111,H1:H3)", 1, "issue #1739 table (VAR.P, row 2 hidden)"),
            ("Sheet1", "E6", "SUBTOTAL(9,D3:D6)", 6, "issue #1739 table (nested SUBTOTAL)"),
         ]),
    dict(id="F25-decimal-leading-sign", issue=1615,
         sources=[IC + "1615"],
         note="The issue table states Excel's #NUM! for a leading sign.",
         inputs={}, cases=[
            ("Sheet1", "A1", '_xlfn.DECIMAL("-10",2)', "#NUM!", "issue #1615 table"),
            ("Sheet1", "A2", "_xlfn.DECIMAL(-5,10)", "#NUM!", "issue #1615 table"),
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
        for sn, cell, formula, expected in fx.get("helpers") or []:
            sh(sn).write_formula(cell, "=" + formula, None, expected)
        for sn, cell, formula, expected, _src in fx["cases"]:
            fmt = datefmt if fx.get("date_format") else None
            sh(sn).write_formula(cell, "=" + formula, fmt, expected)
        for sn, row in fx.get("hidden_rows") or []:
            sh(sn).set_row(row, None, None, {"hidden": True})
    wb.close()


if __name__ == "__main__":
    manifest = []
    for fx in FIXTURES:
        name = fx["id"] + ".xlsx"
        build(fx, os.path.join(HERE, name))
        manifest.append({k: v for k, v in fx.items() if k not in ("inputs", "helpers", "hidden_rows")} | {"file": name})
    with open(os.path.join(HERE, "fixtures.json"), "w") as fh:
        json.dump(manifest, fh, indent=1)
        fh.write("\n")
    print("wrote", len(FIXTURES), "fixtures,", sum(len(f["cases"]) for f in FIXTURES), "cases")
