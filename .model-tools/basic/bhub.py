"""Hub note listing every basic NPV model in Screener order.  -> notes_hub/Basic NPV Models.md"""
import json, os, sys

D = os.path.dirname(os.path.abspath(__file__))
# close date: --close YYYY-MM-DD, else the close the prices in screener_cache.json are from
CLOSE = sys.argv[sys.argv.index("--close") + 1] if "--close" in sys.argv else \
    json.load(open(f"{D}/screener_cache.json"))["close"]
rows = [x for x in json.load(open(f"{D}/tickers.json")) if not x["built"]]


def usd(v):
    return "" if v is None else ("−" if v < 0 else "") + f"\\${abs(v):,.2f}"


def pct(v, sign=True):
    if v is None or isinstance(v, str):
        return v or ""
    return ("−" if v < 0 else "+" if sign else "") + f"{abs(v) * 100:.0f}%"


def mm(v):
    if v is None:
        return ""
    a = abs(v)
    s = f"\\${a / 1000:,.1f}B" if a >= 1000 else f"\\${a:,.0f}M"
    return ("−" if v < 0 else "") + s


dcf, nc, nd, pipe = [], [], [], []
STG = {4: "approved", 3.5: "filed", 3.0: "Ph 3", 2.5: "Ph 2/3", 2.0: "Ph 2", 1.5: "Ph 1/2", 1.0: "Ph 1"}
for x in rows:
    t = x["t"]
    L, C = json.load(open(f"{D}/out/{t}_layout.json")), json.load(open(f"{D}/out/{t}_calc.json"))
    m = C["main"]
    name = x["name"].replace("|", "/") + (" (reviewed: [[Large-Cap Upside Review]])" if L["P"].get("review") else "")
    if L["kind"] == "DCF":
        im = C["implied_m10"]
        dcf.append(f"| [[{t}]] | {name} | {usd(m['C11'])} | {usd(L['price'])} | {pct(m['C12']) if m['C11'] >= 0 else 'below 0'} | "
                   f"{usd(C['scen']['Bear']['ps'])} / {usd(C['scen']['Bull']['ps'])} | {pct(L['P']['m10'], False)} | "
                   f"{pct(im, False) if not isinstance(im, str) else im} |")
    elif L["kind"] == "rNPV":
        p0 = L["programs"][0]
        ip = C.get("implied_peak")
        pipe.append(f"| [[{t}]] | {name} | {usd(m['C11'])} | {usd(L['price'])} | {pct(m['C12']) if (m['C11'] or 0) >= 0 else 'below 0'} | "
                    f"{p0['bucket']} ({STG.get(p0['stage'], p0['stage'])}, {p0['pos'] * 100:.0f}%) | {len(L['programs'])} | "
                    f"{(f'{ip:.1f}×' if isinstance(ip, (int, float)) else ip)} |")
    elif L["kind"] == "Net cash":
        why = "revenue jump" if L.get("flag") else "pre-revenue" if (L["rev_ttm"] or 0) < 5 else "revenue < \\$50M"
        rw = m.get("C18")
        nc.append(f"| [[{t}]] | {name} | {usd(m['C11'])} | {usd(L['price'])} | {pct(m['C12'])} | {mm(m['C9'])} | "
                  f"{f'{rw:.1f} yrs' if isinstance(rw, (int, float)) else (rw if rw == 'net debt' else '')} | {why} |")
    else:
        miss = [k for k, v in (("shares", L["sh"]), ("cash", L["cash"])) if v is None]
        nd.append(f"[[{t}]] ({', '.join(miss)})")
up = [json.load(open(f"{D}/out/{x['t']}_calc.json"))["main"]["C12"] for x in rows
      if json.load(open(f"{D}/out/{x['t']}_layout.json"))["kind"] == "DCF"]
pos = sum(1 for u in up if u > 0)
text = f"""---
Updated: {CLOSE}
---
# Basic NPV Models

Basic valuations for the {len(rows)} Screener Pharma tickers that didn't have a model (everything except LLY, EVMN and the 14 [[PDUFA Calendar]] names). Each workbook in `Stock Models/Pharma/` has every fiscal year of history in SEC XBRL, and each company has a note in `Finance/Companies/`. Prices are the Screener closes of {CLOSE}. No options sheets.

> [!important] How to read these
> - **DCF ({len(dcf)} companies, TTM revenue ≥ \\$50M):** 10-year free-cash-flow DCF built only from reported numbers. Revenue growth starts at the lower of the latest and median recent growth (capped −10% to +35%) and fades to 2.5%; the operating margin moves from TTM to a target (recent median, floored at 20% or gross margin less 15 points); discount rate 8–12.5% by size. {pos} of {len(dcf)} come out above their price.
> - **Pipeline rNPV ({len(pipe)}):** pre-revenue and small-revenue companies with an active pipeline on ClinicalTrials.gov. One program per indication; approval odds from BIO 2011–2020 phase-transition rates; peak sales from an indication benchmark table; company burn until the first launch; 12% discount rate. Filings and approvals were checked by hand for PDUFA dates and launched products. The last column is the peak-sales multiple the price implies (above 1× = the market expects more than the benchmarks).
> - **Net cash ({len(nc)}):** revenue jumps that may be one-time, plus pre-revenue companies with no active registered trials (preclinical, shells, or names that pivoted away from drugs). The value shown is net cash per share.
> - **No data ({len(nd)}):** SEC XBRL lacks a usable share count or cash figure.
> - These are screens, not price targets: no company-specific forecasts or guidance. I'm not a licensed advisor.

## DCF

| Ticker | Company | NPV / share | Price | Upside | Bear / bull | 2035 margin (base) | Price-implied margin |
| --- | --- | --- | --- | --- | --- | --- | --- |
""" + "\n".join(dcf) + f"""

## Pipeline rNPV

| Ticker | Company | rNPV / share | Price | Upside | Lead program (stage, approval odds) | Programs | Price-implied peak multiple |
| --- | --- | --- | --- | --- | --- | --- | --- |
""" + "\n".join(pipe) + f"""

## Net cash

| Ticker | Company | Net cash / share | Price | vs. price | Enterprise value | Runway | Why no DCF |
| --- | --- | --- | --- | --- | --- | --- | --- |
""" + "\n".join(nc) + f"""

## No data

{" · ".join(nd)}

Related: [[PDUFA Calendar]] · [[Stock Model]]
"""
os.makedirs(f"{D}/notes_hub", exist_ok=True)
open(f"{D}/notes_hub/Basic NPV Models.md", "w").write(text)
print(len(dcf), len(pipe), len(nc), len(nd), "positive", pos)
