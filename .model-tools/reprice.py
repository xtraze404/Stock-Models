"""Reprice an installed model workbook without rebuilding it: Main C3 = new price, recalculate every formula and data table
with the `formulas` engine, rewrite the cached values, refresh the Screener price cache in its external link.
usage: reprice.py WORKBOOK.xlsx PRICE OUT.xlsx  (prices for the link cache come from screener_cache.json next to this file)"""
import json, os, re, sys, tempfile, zipfile
from xml.sax.saxutils import escape
import openpyxl, formulas
from openpyxl.worksheet.formula import DataTableFormula
from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl

D = os.path.dirname(os.path.abspath(__file__))
src, price, dst = sys.argv[1], float(sys.argv[2]), sys.argv[3]
wb = openpyxl.load_workbook(src)
wb["Main"]["C3"].value = price
wb._external_links = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, DataTableFormula):
                c.value = None
tmp = tempfile.mkdtemp()
vpath = os.path.join(tmp, "m_v.xlsx")
wb.save(vpath)
X = lambda s, a: f"'[m_v.xlsx]{s.upper()}'!{a}".replace("$", "")
xl = formulas.ExcelModel().loads(vpath).finish()


def val(v):
    x = v.value[0, 0] if hasattr(v, "value") else v
    x = x.item() if hasattr(x, "item") else x
    if x.__class__.__name__ == "XlError" or (isinstance(x, str) and x.startswith("#")):
        return str(x)
    return x


vals = {k: val(v) for k, v in xl.calculate().items()}
zin = zipfile.ZipFile(src)
wbx = zin.read("xl/workbook.xml").decode()
rels = zin.read("xl/_rels/workbook.xml.rels").decode()
rid = {m.group(1): m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels)}
rid.update({m.group(2): m.group(1) for m in re.finditer(r'Target="([^"]+)"[^>]*?Id="(rId\d+)"', rels)})
sheets = {}
for name, r in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
    t = rid[r].lstrip("/")
    sheets["xl/" + t if not t.startswith("xl/") else t] = name.replace("&amp;", "&")

dtk = {}
run = lambda inp, out: val(xl.calculate(inputs=inp, outputs=[out])[out])
for path, sname in sheets.items():
    xml = zin.read(path).decode()
    for m in re.finditer(r'<f t="dataTable"([^>]*)/?>', xml):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        (c0, r0), (c1, r1) = [(ci(re.match(r"[A-Z]+", x).group()), int(re.search(r"\d+", x).group()))
                              for x in a["ref"].split(":")]
        hv = lambda c, r: vals.get(X(sname, f"{cl(c)}{r}"))
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if a.get("dt2D") == "1":
                    inp, out = {X(sname, a["r1"]): hv(c, r0 - 1), X(sname, a["r2"]): hv(c0 - 1, r)}, \
                        X(sname, f"{cl(c0 - 1)}{r0 - 1}")
                elif a.get("dtr") == "1":
                    inp, out = {X(sname, a["r1"]): hv(c, r0 - 1)}, X(sname, f"{cl(c0 - 1)}{r}")
                else:
                    inp, out = {X(sname, a["r1"]): hv(c0 - 1, r)}, X(sname, f"{cl(c)}{r0 - 1}")
                v = run(inp, out)
                vals[X(sname, f"{cl(c)}{r}")] = v
                dtk[X(sname, f"{cl(c)}{r}")] = v
if dtk:      # second pass: cells that read data-table results (e.g. an Options sheet)
    for k, v in xl.calculate(inputs=dtk).items():
        if k not in dtk:
            vals[k] = val(v)
errs = [(k, v) for k, v in vals.items() if isinstance(v, str) and v.startswith("#")]


def vxml(v):
    if v is None or (isinstance(v, str) and v == ""):
        return None
    if isinstance(v, bool):
        return ' t="b"', "1" if v else "0"
    if isinstance(v, (int, float)):
        return "", repr(float(v)) if isinstance(v, float) else str(v)
    s = str(v)
    return (' t="e"', s) if s.startswith("#") else (' t="str"', escape(s))


out = {}
for path, sname in sheets.items():
    xml = zin.read(path).decode()
    dtref = set()
    for m in re.finditer(r'<f t="dataTable" ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml):
        for r in range(int(m.group(2)), int(m.group(4)) + 1):
            for c in range(ci(m.group(1)), ci(m.group(3)) + 1):
                dtref.add(f"{cl(c)}{r}")

    def fix(m):
        attrs, body = m.group(1), m.group(3) or ""
        addr = re.search(r'r="([A-Z]+\d+)"', attrs).group(1)
        is_f = "<f" in body
        if not (is_f or addr in dtref):
            return m.group(0)
        vx = vxml(vals.get(X(sname, addr)))
        if vx is None:
            return m.group(0)
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
        fpart = re.search(r"<f[^>]*/>|<f[^>]*>.*?</f>", body, re.S)
        return f"<c{attrs}{vx[0]}>{fpart.group(0) if fpart else ''}<v>{vx[1]}</v></c>"

    out[path] = re.sub(r"<c( [^>]*?)(/>|>(.*?)</c>)", lambda m: fix(m) if m.group(2) != "/>" or
                       re.search(r'r="([A-Z]+\d+)"', m.group(1)).group(1) in dtref else m.group(0), xml, flags=re.S)
# Screener price cache inside the external link (column F of the Pharma sheet)
SC = {i: p for i, t, p in json.load(open(f"{D}/basic/screener_cache.json"))["rows"] if isinstance(p, (int, float))}
for name in zin.namelist():
    m = re.fullmatch(r"xl/externalLinks/(externalLink\d+)\.xml", name)
    if not m:
        continue
    x = zin.read(name).decode()
    x = re.sub(r'(<(?:\w+:)?cell r="F(\d+)"[^>]*>\s*<(?:\w+:)?v>)([^<]*)(</(?:\w+:)?v>)',
               lambda mm: mm.group(1) + (repr(float(SC[int(mm.group(2))])) if int(mm.group(2)) in SC else mm.group(3)) + mm.group(4), x)
    out[name] = x
    rid_ = re.search(r'<externalBook[^>]*r:id="([^"]+)"', x).group(1)
    rn = f"xl/externalLinks/_rels/{m.group(1)}.xml.rels"
    out[rn] = re.sub(r'Id="[^"]+"', f'Id="{rid_}"', zin.read(rn).decode())
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zin.infolist():
        data = out[it.filename].encode() if it.filename in out else zin.read(it.filename)
        zo.writestr(it, data, compress_type=zipfile.ZIP_DEFLATED)
g = lambda a: vals.get(X("Main", a))
print(json.dumps({"price": price, "C11": g("C11"), "C12": g("C12"), "cells": len(vals), "errors": len(errs),
                  "first_errors": errs[:3]}, default=str))
