"""Download SEC companyfacts + submissions (gzipped) for every ticker to build."""
import json, gzip, os, time, urllib.request, threading
from concurrent.futures import ThreadPoolExecutor
UA = {"User-Agent": "StockModel research research@example.com", "Accept-Encoding": "gzip"}
D = os.path.dirname(os.path.abspath(__file__))
os.makedirs(f"{D}/raw", exist_ok=True)
rows = [x for x in json.load(open(f"{D}/tickers.json")) if not x["built"]]
lock = threading.Lock(); last = [0.0]
def get(url):
    for k in range(5):
        with lock:
            w = last[0] + 0.12 - time.time()
            if w > 0: time.sleep(w)
            last[0] = time.time()
        try:
            r = urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=120)
            b = r.read()
            return b if r.headers.get("Content-Encoding") != "gzip" else gzip.decompress(b)
        except urllib.error.HTTPError as e:
            if e.code == 404: return None
            time.sleep(2 + 3 * k)
        except Exception:
            time.sleep(2 + 3 * k)
    return None
def job(x):
    t, cik = x["t"], int(x["cik"])
    out = []
    for kind, url in (("facts", f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"),
                      ("subs", f"https://data.sec.gov/submissions/CIK{cik:010d}.json")):
        p = f"{D}/raw/{t}_{kind}.json.gz"
        if os.path.exists(p): continue
        b = get(url)
        if b is None: out.append(f"{kind} missing"); continue
        with gzip.open(p, "wb") as f: f.write(b)
    return t, out
n = 0
with ThreadPoolExecutor(6) as ex:
    for t, out in ex.map(job, rows):
        n += 1
        if out: print(t, out, flush=True)
        if n % 100 == 0: print("done", n, flush=True)
print("finished", n)
