"""Regular-session closes right after the bell: Nasdaq quote feed ('Closed at <date> 4:00 PM ET'), Cboe close as fallback.
usage: closes_live.py 'Oct 20, 2026' 2026-10-20 OUT.json TICKER..."""
import json, sys, time, urllib.request, concurrent.futures as cf
stamp, iso, out_p, tick = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4:]
H = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
     "Accept": "application/json"}
def nasdaq(sym):
    for cls in ("stocks", "etf"):
        for attempt in range(3):
            try:
                u = f"https://api.nasdaq.com/api/quote/{sym}/info?assetclass={cls}"
                d = json.load(urllib.request.urlopen(urllib.request.Request(u, headers=H), timeout=20)).get("data")
                if not d: break
                for k in ("secondaryData", "primaryData"):
                    s = d.get(k) or {}
                    ts = (s.get("lastTradeTimestamp") or "").strip()
                    ok = f"Closed at {stamp}" in ts or (k == "primaryData" and ts == stamp and d.get("marketStatus") == "Closed")
                    if ok and s.get("lastSalePrice"):
                        return float(s["lastSalePrice"].replace("$", "").replace(",", "")), "nasdaq"
                break
            except Exception:
                time.sleep(1.5 * (attempt + 1))
    return None, None
def cboe(sym):
    try:
        d = json.load(urllib.request.urlopen(urllib.request.Request(
            f"https://cdn.cboe.com/api/global/delayed_quotes/options/{sym}.json", headers=H), timeout=20))["data"]
        if (d.get("last_trade_time") or "")[:10] == iso and d.get("close"):
            return float(d["close"]), "cboe"
    except Exception:
        pass
    return None, None
def get(t):
    sym = {"QNCX": "IRLA"}.get(t, t)
    p, src = nasdaq(sym)
    if p is None:
        p, src = cboe(sym)
    return t, p, src
out = {}
with cf.ThreadPoolExecutor(3) as ex:
    for t, p, src in ex.map(get, tick):
        out[t] = {"px": p, "src": src}
json.dump(out, open(out_p, "w"), indent=0)
from collections import Counter
print(Counter(v["src"] for v in out.values()), "| missing:", [t for t, v in out.items() if v["px"] is None])
