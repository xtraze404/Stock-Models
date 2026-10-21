"""Next earnings date per ticker from Nasdaq (Zacks).  usage: earnings.py OUT.json TICKER..."""
import json, re, sys, time, datetime as dt, urllib.request, concurrent.futures as cf
out_p, tick = sys.argv[1], sys.argv[2:]
H = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
     "Accept": "application/json"}
def get(t):
    sym = {"QNCX": "IRLA"}.get(t, t)
    for attempt in range(3):
        try:
            d = json.load(urllib.request.urlopen(urllib.request.Request(
                f"https://api.nasdaq.com/api/analyst/{sym}/earnings-date", headers=H), timeout=20)).get("data") or {}
            m = re.search(r":\s*([A-Z][a-z]{2} \d{1,2}, \d{4})", d.get("announcement") or "")
            if not m:
                return t, None
            date = dt.datetime.strptime(m.group(1), "%b %d, %Y").date().isoformat()
            rt = d.get("reportText") or ""
            when = "before open" if "before market open" in rt else "after close" if "after market close" in rt else ""
            return t, {"date": date, "status": "estimated" if "estimated to report" in rt else "expected", "when": when}
        except Exception:
            time.sleep(1.5 * (attempt + 1))
    return t, "error"
out = {}
with cf.ThreadPoolExecutor(3) as ex:
    for t, v in ex.map(get, tick):
        out[t] = v
json.dump(out, open(out_p, "w"), indent=0)
print("dates", sum(isinstance(v, dict) for v in out.values()), "none", sum(v is None for v in out.values()),
      "errors", [t for t, v in out.items() if v == "error"])
