"""Read formula text, formula attributes, type, and cached value from OOXML."""
import posixpath
import zipfile
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
    out = {}
    for s in wb.find("m:sheets", NS).findall("m:sheet", NS):
        t = rmap[s.get("{%s}id" % NS["r"])]
        t = t.lstrip("/") if t.startswith("/") else posixpath.join("xl", t)
        out[s.get("name")] = t
    return out


def shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])) for si in root.findall("m:si", NS)]


def cells(path, wanted):
    """wanted: list of (sheet, ref). Returns {(sheet, ref): dict(t, f, fattr, v, cm)}."""
    z = zipfile.ZipFile(path)
    sp = sheet_paths(z)
    ss = shared_strings(z)
    out = {}
    by_sheet = {}
    for s, r in wanted:
        by_sheet.setdefault(s, set()).add(r)
    for s, refs in by_sheet.items():
        if s not in sp:
            for r in refs:
                out[(s, r)] = None
            continue
        root = ET.fromstring(z.read(sp[s]))
        found = {}
        for c in root.iter("{%s}c" % NS["m"]):
            if c.get("r") in refs:
                f = c.find("m:f", NS)
                v = c.find("m:v", NS)
                t = c.get("t")
                val = v.text if v is not None else None
                if t == "s" and val is not None:
                    val = ss[int(val)]
                found[c.get("r")] = dict(
                    t=t,
                    f=(f.text if f is not None else None),
                    fattr=(dict(f.attrib) if f is not None else None),
                    v=val,
                    cm=c.get("cm"),
                )
        for r in refs:
            out[(s, r)] = found.get(r)
    return out


def parts(path):
    return sorted(zipfile.ZipFile(path).namelist())
