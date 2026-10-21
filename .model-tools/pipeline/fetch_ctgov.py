"""ClinicalTrials.gov v2: every interventional study naming the company as sponsor or collaborator.
usage: fetch_ctgov.py TICKERS_FILE   -> raw/<T>.json  (skips tickers already fetched)"""
import json, os, re, sys, time, urllib.parse, urllib.request
from concurrent.futures import ThreadPoolExecutor
D = os.path.dirname(os.path.abspath(__file__))
TK = {x["t"]: x for x in json.load(open(f"{D}/../basic/tickers.json"))}
FIELDS = ("NCTId,BriefTitle,Phase,Condition,InterventionName,InterventionType,OverallStatus,StudyType,StartDate,"
          "PrimaryCompletionDate,CompletionDate,LeadSponsorName,LeadSponsorClass,CollaboratorName,EnrollmentCount,"
          "LastUpdatePostDate")
SUFFIX = r"(,?\s+(inc\.?|incorporated|corp\.?|corporation|ltd\.?|limited|plc|n\.?v\.?|s\.?a\.?|ag|a/s|se|ab|llc|co\.?|holdings?|group|international|common stock|ordinary shares|american depositary shares.*))+$"
def clean(name):
    n = re.sub(r"\s*\(.*?\)", "", name or "").strip()
    for _ in range(3):
        n = re.sub(SUFFIX, "", n, flags=re.I).strip(" ,.")
    return n
def get(url):
    for k in range(4):
        try:
            return json.load(urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=60))
        except Exception:
            time.sleep(2 + 3 * k)
    return None
def job(t):
    out = f"{D}/raw/{t}.json"
    if os.path.exists(out):
        return t, "cached"
    q = clean(TK[t]["name"])
    studies, tok = [], None
    for _ in range(10):
        u = (f"https://clinicaltrials.gov/api/v2/studies?query.spons={urllib.parse.quote(q)}&pageSize=200&fields={FIELDS}"
             + (f"&pageToken={tok}" if tok else ""))
        d = get(u)
        if d is None:
            return t, "error"
        studies += d.get("studies", [])
        tok = d.get("nextPageToken")
        if not tok:
            break
        time.sleep(0.3)
    json.dump({"query": q, "studies": studies}, open(out, "w"))
    return t, len(studies)
ts = [l.strip() for l in open(sys.argv[1]) if l.strip()]
with ThreadPoolExecutor(4) as ex:
    res = list(ex.map(job, ts))
print("fetched", len(res), "| zero studies:", [t for t, n in res if n == 0][:40], "| errors:", [t for t, n in res if n == "error"])
