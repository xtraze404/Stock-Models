"""openpyxl renumbers each external link's relationship Id; make every rels Id match its <externalBook r:id>."""
import re, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src); fixed = {}; n_fix = 0
for name in zin.namelist():
    m = re.fullmatch(r"xl/externalLinks/(externalLink\d+)\.xml", name)
    if not m: continue
    rid = re.search(r'<externalBook r:id="([^"]+)"', zin.read(name).decode()).group(1)
    rn = f"xl/externalLinks/_rels/{m.group(1)}.xml.rels"
    rels = zin.read(rn).decode()
    new = re.sub(r'Id="[^"]+"', f'Id="{rid}"', rels)
    n_fix += new != rels; fixed[rn] = new.encode()
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
    for it in zin.infolist():
        z.writestr(it, fixed.get(it.filename, zin.read(it.filename)))
print("links:", len(fixed), "ids fixed:", n_fix)
