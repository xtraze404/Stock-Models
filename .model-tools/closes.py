"""Official daily closes from Nasdaq's historical API.  usage: closes.py MM/DD/YYYY OUT.json TICKER..."""
import json, sys, time, urllib.request, concurrent.futures as cf
day, out_p, tick = sys.argv[1], sys.argv[2], sys.argv[3:]
H = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
     "Accept": "application/json"}
mm, dd, yy = day.split("/")
iso = f"{yy}-{mm}-{dd}"
def get(t):
    sym = {"QNCX": "IRLA"}.get(t, t)          # QNCX trades as IRLA since 2026-10-09 (IRulya Therapeutics)
    err = None
    for attempt in range(4):
        try:
            u = f"https://api.nasdaq.com/api/quote/{sym}/historical?assetclass=stocks&fromdate={iso}&todate={iso}&limit=5"
            d = json.load(urllib.request.urlopen(urllib.request.Request(u, headers=H), timeout=20)).get("data")
            rows = ((d or {}).get("tradesTable") or {}).get("rows") or []
            px = {r["date"]: float(r["close"].replace("$", "").replace(",", "")) for r in rows}
            return t, px.get(day), None
        except Exception as e:
            time.sleep(1.5 * (attempt + 1)); err = str(e)
    return t, None, err
out = {}
with cf.ThreadPoolExecutor(6) as ex:
    for t, p, err in ex.map(get, tick):
        out[t] = p
json.dump(out, open(out_p, "w"), indent=0)
print("closes", sum(v is not None for v in out.values()), "of", len(out), "| missing:", [t for t, v in out.items() if v is None])
