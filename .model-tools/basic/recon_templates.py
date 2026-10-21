"""Rebuild templates.json (template Main C4/C7/C8 + their SEC source comments) from the installed workbooks and notes.
Fields the template left blank (the 'Left blank: ...' comments) stay None so bbuild.py re-derives them the same way."""
import json, re, openpyxl
D = "/Users/eric/Documents/Stock Models/.model-tools/basic"
V = "/Users/eric/Documents/Vault/Finance/Companies"
todo = [l.strip() for l in open(f"{D}/todo.txt") if l.strip()]
out = {}
for t in todo:
    m = openpyxl.load_workbook(f"/Users/eric/Documents/Stock Models/Pharma/{t}.xlsx")["Main"]
    note = open(f"{V}/{t}.md").read()
    com = {}
    for lab, key in (("Shares", "C4c"), ("Cash", "C7c"), ("Debt", "C8c")):
        mm = re.search(rf"^- \*\*{lab} source:\*\* (.*)$", note, re.M)
        com[key] = mm.group(1).replace("\\$", "$") if mm else None
    d = {"C3": None}
    for a in ("C4", "C7", "C8"):
        v = m[a].value
        c = com.get(a + "c") or ""
        d[a] = None if c.startswith("Left blank") else v
    d.update(com)
    out[t] = d
json.dump(out, open(f"{D}/templates.json", "w"), indent=0, default=str)
print(len(out), "| None C4:", sum(v["C4"] is None for v in out.values()), "None C7:", sum(v["C7"] is None for v in out.values()),
      "| no comments:", sum(v["C4c"] is None for v in out.values()))
