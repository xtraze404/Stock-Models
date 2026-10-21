"""Hub note listing every basic NPV model in Screener order.  -> notes_hub/Basic NPV Models.md"""
import json, os

D = os.path.dirname(os.path.abspath(__file__))
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


dcf, nc, nd = [], [], []
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
Updated: 2026-10-20
---
# Basic NPV Models

Basic valuations for the {len(rows)} Screener Pharma tickers that didn't have a model (everything except LLY, EVMN and the 14 [[PDUFA Calendar]] names). Each workbook in `Stock Models/Pharma/` has every fiscal year of history in SEC XBRL, and each company has a note in `Finance/Companies/`. Prices are the Screener closes of 2026-10-20. No options sheets.

> [!important] How to read these
> - **DCF ({len(dcf)} companies, TTM revenue ≥ \\$50M):** 10-year free-cash-flow DCF built only from reported numbers. Revenue growth starts at the lower of the latest and median recent growth (capped −10% to +35%) and fades to 2.5%; the operating margin moves from TTM to a target (recent median, floored at 20% or gross margin less 15 points); discount rate 8–12.5% by size. {pos} of {len(dcf)} come out above their price.
> - **Net cash ({len(nc)}):** pre-revenue companies, those under \\$50M of revenue, and revenue jumps that may be one-time (licence or milestone income, or a launch too new to extrapolate). The value shown is net cash per share; enterprise value is what the market pays for the pipeline.
> - **No data ({len(nd)}):** SEC XBRL lacks a usable share count or cash figure.
> - These are screens, not price targets: no product forecasts, patent expiries, pipeline or guidance. I'm not a licensed advisor.

## DCF

| Ticker | Company | NPV / share | Price | Upside | Bear / bull | 2035 margin (base) | Price-implied margin |
| --- | --- | --- | --- | --- | --- | --- | --- |
""" + "\n".join(dcf) + f"""

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
print(len(dcf), len(nc), len(nd), "positive", pos)
