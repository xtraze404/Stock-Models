"""Compact history from SEC companyfacts: every fiscal year available, TTM, latest balance-sheet date, currency.
usage: extract.py [TICKERS...]  -> hist/<T>.json  (all tickers to build when none given)"""
import datetime as dt, gzip, json, os, sys
from collections import defaultdict

D = os.path.dirname(os.path.abspath(__file__))
os.makedirs(f"{D}/hist", exist_ok=True)
G, I = "us-gaap", "ifrs-full"
FXR = json.load(open(f"{D}/raw/fx.json"))["rates"]
FXR["USD"] = 1.0
# per field: list of (namespace, concept); "max" fields take the largest value across concepts for a year
DUR = {
    "rev": ("max", [(G, "Revenues"), (G, "RevenueFromContractWithCustomerExcludingAssessedTax"),
                    (G, "RevenueFromContractWithCustomerIncludingAssessedTax"), (G, "SalesRevenueNet"),
                    (G, "SalesRevenueGoodsNet"), (I, "Revenue"), (I, "RevenueFromContractsWithCustomers"), (I, "RevenueFromSaleOfGoods"),
                    (I, "RevenueAndOperatingIncome")]),
    "cogs": ("max", [(G, "CostOfGoodsAndServicesSold"), (G, "CostOfRevenue"), (G, "CostOfGoodsSold"),
                       (G, "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"),
                       (G, "CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization"), (I, "CostOfSales")]),
    "rd": ("max", [(G, "ResearchAndDevelopmentExpense"), (G, "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"),
                     (I, "ResearchAndDevelopmentExpense")]),
    "sga": ("max", [(G, "SellingGeneralAndAdministrativeExpense"), (G, "GeneralAndAdministrativeExpense"),
                      (I, "SellingGeneralAndAdministrativeExpense"), (I, "AdministrativeExpense")]),
    "opex": ("first", [(G, "OperatingExpenses"), (I, "OperatingExpense")]),
    "costs": ("first", [(G, "CostsAndExpenses")]),
    "oi": ("first", [(G, "OperatingIncomeLoss"), (I, "ProfitLossFromOperatingActivities")]),
    "pretax": ("first", [(G, "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
                         (G, "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"),
                         (G, "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic"), (I, "ProfitLossBeforeTax")]),
    "tax": ("first", [(G, "IncomeTaxExpenseBenefit"), (I, "IncomeTaxExpenseContinuingOperations")]),
    "ni": ("first", [(G, "NetIncomeLoss"), (G, "NetIncomeLossAvailableToCommonStockholdersBasic"), (G, "ProfitLoss"),
                     (I, "ProfitLossAttributableToOwnersOfParent"), (I, "ProfitLoss")]),
    "cfo": ("first", [(G, "NetCashProvidedByUsedInOperatingActivities"),
                      (G, "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"),
                      (I, "CashFlowsFromUsedInOperatingActivities")]),
    "sh": ("first", [(G, "WeightedAverageNumberOfDilutedSharesOutstanding"),
                     (G, "WeightedAverageNumberOfShareOutstandingBasicAndDiluted"),
                     (G, "WeightedAverageNumberOfSharesOutstandingBasic"),
                     (I, "AdjustedWeightedAverageShares"), (I, "WeightedAverageShares")]),
}
BS = [(G, "Assets"), (I, "Assets")]


def facts_of(t):
    return json.load(gzip.open(f"{D}/raw/{t}_facts.json.gz"))["facts"]


def units(facts, ns, c):
    return facts.get(ns, {}).get(c, {}).get("units", {})


def currency(facts):
    """Reporting currency = the unit of the most recent revenue / net income / assets fact."""
    best = None
    for ns, c in DUR["rev"][1] + DUR["ni"][1] + BS:
        for u, xs in units(facts, ns, c).items():
            if len(u) == 3 and u.isupper():
                for x in xs:
                    k = (x["end"], x["filed"])
                    if best is None or k > best[0]:
                        best = (k, u)
    return best[1] if best else "USD"


def fylabel(end):
    e = dt.date.fromisoformat(end)
    return e.year - 1 if e.month == 1 and e.day <= 10 else e.year


def periods(xs):
    """{(start, end): fact} keeping the latest-filed fact for each period."""
    out = {}
    for x in xs:
        if "start" not in x:
            continue
        k = (x["start"], x["end"])
        if k not in out or x["filed"] > out[k]["filed"]:
            out[k] = x
    return out


def days(k):
    return (dt.date.fromisoformat(k[1]) - dt.date.fromisoformat(k[0])).days


def usd_facts(facts, ns, c, unit):
    """Facts for a concept: shares as is; money in every currency unit converted to USD at the Oct 16 rate."""
    us = units(facts, ns, c)
    if unit == "shares":
        return us.get("shares", [])
    out = []
    for u, xs in us.items():
        if u in FXR:
            out += [dict(x, val=x["val"] / FXR[u], _u=u) for x in xs]
    return out


def field(facts, key, unit):
    mode, cands = DUR[key]
    ann, ytd = {}, {}            # ann: label -> (end, val, concept); ytd: (start, end) -> val
    for ns, c in cands:
        xs = usd_facts(facts, ns, c, unit)
        if not xs:
            continue
        ps = periods(xs)
        a = {}
        for k, x in ps.items():
            if 350 <= days(k) <= 380:
                lab = fylabel(k[1])
                if lab not in a or k[1] > a[lab][0]:
                    a[lab] = (k[1], x["val"], c, k[0])
        for lab, v in a.items():
            if lab not in ann or (mode == "max" and v[1] > ann[lab][1]):
                ann[lab] = v
        for k, x in ps.items():
            if 80 <= days(k) <= 290:
                if k not in ytd or (mode == "max" and x["val"] > ytd[k]):
                    ytd[k] = x["val"]
    return ann, ytd


def ttm(ann, ytd):
    """Latest FY plus the year-to-date change after it: (end, value, prior-year TTM)."""
    if not ann:
        return None
    lab = max(ann)
    fy_end, fy_val, _, fy_start = ann[lab]
    prev = ann.get(lab - 1)
    best = None
    for (s, e), v in ytd.items():
        if s > fy_end and e > fy_end:                  # YTD in the new fiscal year
            ds = (dt.date.fromisoformat(s) - dt.date.fromisoformat(fy_end)).days
            if ds > 5:
                continue
            # matching prior-year YTD: same length, one year earlier
            pe = (dt.date.fromisoformat(e) - dt.timedelta(days=365)).isoformat()
            for (s2, e2), v2 in ytd.items():
                if abs((dt.date.fromisoformat(e2) - dt.date.fromisoformat(pe)).days) <= 7 and \
                        abs(days((s2, e2)) - days((s, e))) <= 7:
                    if best is None or e > best[0]:
                        best = (e, v, v2)
    if best:
        cur = fy_val + best[1] - best[2]
        prior = None
        if prev:
            # prior TTM = prior FY + prior YTD - YTD two years back (rarely available) -> use FY growth fallback
            prior = None
        return {"end": best[0], "val": cur, "ytd": best[1], "ytd_prior": best[2], "fy": fy_val, "fy_end": fy_end}
    return {"end": fy_end, "val": fy_val, "ytd": None, "ytd_prior": None, "fy": fy_val, "fy_end": fy_end}


def latest_bs(facts):
    best = None
    for ns, c in BS:
        for u, xs in units(facts, ns, c).items():
            for x in xs:
                if "start" not in x and (best is None or x["end"] > best):
                    best = x["end"]
    return best


def run(t):
    f = facts_of(t)
    cur = currency(f)
    out = {"t": t, "ccy": cur, "bs_end": latest_bs(f), "fy": {}, "ttm": {}, "concepts": {}}
    for key in DUR:
        unit = "shares" if key == "sh" else "USD"          # money already converted to USD per fact
        ann, ytd = field(f, key, unit)
        for lab, (end, val, c, start) in ann.items():
            out["fy"].setdefault(str(lab), {"end": end})[key] = val / 1e6
        out["concepts"][key] = sorted({v[2] for v in ann.values()})
        if key in ("rev", "oi", "cfo", "ni", "rd", "sga", "cogs"):
            tt = ttm(ann, ytd)
            if tt:
                out["ttm"][key] = {k: (v / 1e6 if isinstance(v, (int, float)) else v) for k, v in tt.items()}
    assets = {}
    for ns, c in BS:
        for x in usd_facts(f, ns, c, "USD"):
            if "start" not in x and (x["end"] not in assets or x["filed"] > assets[x["end"]][1]):
                assets[x["end"]] = (x["val"] / 1e6, x["filed"])
    for y, r in out["fy"].items():
        if r["end"] in assets:
            r["assets"] = assets[r["end"]][0]
    out["fy"] = dict(sorted(out["fy"].items()))
    json.dump(out, open(f"{D}/hist/{t}.json", "w"), indent=0)
    return out


if __name__ == "__main__":
    ts = sys.argv[1:] or [x["t"] for x in json.load(open(f"{D}/tickers.json")) if not x["built"]]
    for t in ts:
        try:
            o = run(t)
            if len(ts) <= 8:
                print(t, o["ccy"], o["bs_end"], "years", list(o["fy"])[:1], "..", list(o["fy"])[-1:], o["concepts"])
                for y, r in list(o["fy"].items())[-4:]:
                    print("  ", y, {k: round(v, 1) if isinstance(v, float) else v for k, v in r.items()})
                print("   TTM", {k: round(v["val"], 1) for k, v in o["ttm"].items()}, o["ttm"].get("rev", {}).get("end"))
        except FileNotFoundError:
            print(t, "no facts file")
        except Exception as e:
            print(t, "ERROR", type(e).__name__, e)
