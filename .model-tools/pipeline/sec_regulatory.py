"""Regulatory stage from the latest 10-Q / 10-K: PDUFA dates (NDA/BLA under review) and FDA approvals.
usage: sec_regulatory.py [TICKERS...]  -> sec/<T>.json with matched sentences"""
import gzip, html, json, os, re, sys, time, threading, urllib.request
from concurrent.futures import ThreadPoolExecutor

D = os.path.dirname(os.path.abspath(__file__))
UA = {"User-Agent": "StockModel research research@example.com"}
lock, last = threading.Lock(), [0.0]


def get(url):
    for k in range(4):
        with lock:
            w = last[0] + 0.13 - time.time()
            if w > 0:
                time.sleep(w)
            last[0] = time.time()
        try:
            return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=90).read().decode("utf-8", "ignore")
        except Exception:
            time.sleep(2 + 3 * k)
    return None


def text_of(raw):
    raw = re.sub(r"<ix:header>.*?</ix:header>", " ", raw, flags=re.S | re.I)
    raw = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", raw, flags=re.S | re.I)
    x = html.unescape(re.sub(r"<[^>]+>", " ", raw))
    return re.sub(r"\s+", " ", x)


PDUFA = re.compile(r"[^.]{0,300}\b(PDUFA|Prescription Drug User Fee Act)\b[^.]{0,300}\.", re.I)
APPROVED = re.compile(r"[^.]{0,250}\b(FDA|Food and Drug Administration|European Commission|EMA)\b[^.]{0,60}\b(approved|granted (full |accelerated )?approval|approval of)\b[^.]{0,250}\.", re.I)
HEDGE = re.compile(r"\b(if|whether|potential(ly)?|may|might|could|seek|obtain|expect|plan|intend|anticipate|prior to|unless|until|would|should|risk)\b", re.I)
DATE = re.compile(r"(January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, 20\d\d")


def job(t):
    out = f"{D}/sec/{t}.json"
    if os.path.exists(out):
        return t, "cached"
    try:
        s = json.load(gzip.open(f"{D}/../basic/raw/{t}_subs.json.gz"))
    except FileNotFoundError:
        return t, "no subs"
    r = s["filings"]["recent"]
    cik = int(s["cik"])
    idx = next((i for i, f in enumerate(r["form"]) if f in ("10-Q", "10-K", "20-F", "40-F")), None)
    if idx is None:
        json.dump({"form": None}, open(out, "w"))
        return t, "no periodic report"
    acc = r["accessionNumber"][idx].replace("-", "")
    url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{r['primaryDocument'][idx]}"
    raw = get(url)
    if raw is None:
        return t, "error"
    x = text_of(raw)
    pd = [m.group(0).strip() for m in PDUFA.finditer(x)][:15]
    ap = [m.group(0).strip() for m in APPROVED.finditer(x) if not HEDGE.search(m.group(0))][:15]
    json.dump({"form": r["form"][idx], "date": r["filingDate"][idx], "url": url, "pdufa": pd, "approved": ap,
               "pdufa_dates": sorted({d for sent in pd for d in DATE.findall(sent) and [m.group(0) for m in DATE.finditer(sent)]})},
              open(out, "w"), indent=1)
    return t, f"pdufa {len(pd)} approved {len(ap)}"


if __name__ == "__main__":
    os.makedirs(f"{D}/sec", exist_ok=True)
    ts = sys.argv[1:] or [l.strip() for l in open(f"{D}/universe.txt") if l.strip()]
    with ThreadPoolExecutor(6) as ex:
        res = list(ex.map(job, ts))
    from collections import Counter
    print(Counter(v.split()[0] if v != "cached" else v for _, v in res))
