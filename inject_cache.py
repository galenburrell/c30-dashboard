#!/usr/bin/env python3
"""
Write cached formula results into an openpyxl-built workbook without letting LibreOffice rewrite it.

openpyxl saves formulas with no cached values (previews and phone viewers then show blanks), and a
LibreOffice round-trip strips some chart styling. This copies the values LibreOffice computed
(recalculated copy) into the <v> elements of the original file's formula cells.

Usage: python3 inject_cache.py original.xlsx recalculated.xlsx output.xlsx
"""
import sys, zipfile, datetime, re
import openpyxl
from lxml import etree

orig, recalc, outp = sys.argv[1:4]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EPOCH = datetime.datetime(1899, 12, 30)

vals = {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
        for ws in openpyxl.load_workbook(recalc, data_only=True, read_only=False).worksheets}

zin = zipfile.ZipFile(orig)
wbx = etree.fromstring(zin.read("xl/workbook.xml"))
rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
rid2target = {r.get("Id"): r.get("Target") for r in rels}
sheet_paths = {}
for s in wbx.find(f"{{{NS}}}sheets"):
    t = rid2target[s.get(f"{{{RNS}}}id")]
    t = t.lstrip("/")
    sheet_paths["xl/" + t if not t.startswith("xl/") else t] = s.get("name")

stats = {"filled": 0, "empty": 0}
def patch(xml, sheetname):
    root = etree.fromstring(xml)
    sv = vals.get(sheetname, {})
    for c in root.iter(f"{{{NS}}}c"):
        f = c.find(f"{{{NS}}}f")
        if f is None: continue
        v = c.find(f"{{{NS}}}v")
        if v is None:
            v = etree.SubElement(c, f"{{{NS}}}v")
        val = sv.get(c.get("r"))
        if val is None or val == "":
            c.set("t", "str"); v.text = ""; stats["empty"] += 1; continue
        if isinstance(val, bool):
            c.set("t", "b"); v.text = "1" if val else "0"
        elif isinstance(val, (int, float)):
            c.attrib.pop("t", None); v.text = repr(float(val)) if isinstance(val, float) else str(val)
        elif isinstance(val, datetime.datetime):
            c.attrib.pop("t", None); v.text = repr((val - EPOCH).total_seconds() / 86400)
        elif isinstance(val, datetime.date):
            c.attrib.pop("t", None); v.text = str((datetime.datetime.combine(val, datetime.time()) - EPOCH).days)
        else:
            c.set("t", "str"); v.text = str(val)
        stats["filled"] += 1
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

with zipfile.ZipFile(outp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in sheet_paths:
            data = patch(data, sheet_paths[item.filename])
        zout.writestr(item, data)
print(f"cached values written: {stats['filled']} values, {stats['empty']} blank results")
