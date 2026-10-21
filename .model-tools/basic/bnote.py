"""Obsidian note for a basic NPV model.  usage: bnote.py TICKER [TICKER...]  -> notes/<T>.md (merged with an existing
note when one is in the vault: its frontmatter and text are kept, the model section is appended)."""
import gzip, json, os, re, sys

D = os.path.dirname(os.path.abspath(__file__))
VAULT = os.path.join(os.environ.get("SM_VAULT") or os.path.normpath(f"{D}/../../../Vault"), "Finance", "Companies")
ROWS = {x["t"]: x for x in json.load(open(f"{D}/tickers.json"))}
TPL = json.load(open(f"{D}/templates.json"))


def close_date(argv):
    """--close YYYY-MM-DD (removed from argv), else the close the prices in screener_cache.json are from."""
    if "--close" in argv:
        i = argv.index("--close")
        c = argv[i + 1]
        del argv[i:i + 2]
        return c
    return json.load(open(f"{D}/screener_cache.json"))["close"]


TODAY = close_date(sys.argv)
STG = {4: "Approved", 3.5: "Filed (NDA/BLA)", 3.0: "Phase 3", 2.5: "Phase 2/3", 2.0: "Phase 2", 1.5: "Phase 1/2", 1.0: "Phase 1"}          # price close date; also the notes' "Updated" date
USER_NOTES = {"MRK", "LENZ", "MNKD", "SPRB"}   # existed in the vault before the basic models
REV = json.load(open(f"{D}/review.json")) if os.path.exists(f"{D}/review.json") else {}


def esc(s):
    return s           # every body line is $-escaped once in write()


def money(v, dec=0):
    """$mm -> '$1,234M' or '$12.3B'"""
    if v is None:
        return "n/a"
    a = abs(v)
    s = f"${a / 1000:,.1f}B" if a >= 10000 else f"${a:,.{dec}f}M"
    return ("−" if v < 0 else "") + s


def usd(v):
    if v is None:
        return "n/a"
    return ("−" if v < 0 else "") + f"${abs(v):,.2f}"


def pct(v, dec=0, sign=False):
    if v is None or isinstance(v, str):
        return "n/a"
    s = f"{abs(v) * 100:.{dec}f}%"
    if v < 0:
        return "−" + s
    return ("+" + s) if sign else s


def srcl(src, L):
    if src == "template":
        return f"SEC XBRL, {L.get('bs_end') or 'latest filing'}"
    if src == "n/a":
        return "not in SEC XBRL"
    return f"SEC XBRL, {src}"


def filings(t, cik):
    try:
        s = json.load(gzip.open(f"{D}/raw/{t}_subs.json.gz"))
    except FileNotFoundError:
        return {}, []
    r = s["filings"]["recent"]
    out = []
    for kinds, label in ((("10-K", "20-F", "40-F"), "annual"), (("10-Q",), "quarterly")):
        for i, f in enumerate(r["form"]):
            if f in kinds:
                acc = r["accessionNumber"][i].replace("-", "")
                url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc}/{r['primaryDocument'][i]}"
                per = r["reportDate"][i] or r["filingDate"][i]
                out.append((f"{f} ({per})", url))
                break
    return s, out


def note(t):
    R, L, C = ROWS[t], json.load(open(f"{D}/out/{t}_layout.json")), json.load(open(f"{D}/out/{t}_calc.json"))
    cik = int(R["cik"])
    subs, fl = filings(t, cik)
    name = R["name"]
    sic = f"{subs.get('sic')} - {subs.get('sicDescription')}" if subs.get("sic") else None
    m, P, K = C["main"], L["P"], L["kind"]
    price = L["price"]
    years = [int(y) for y in L["years"]]
    hist = {int(y): h for y, h in L["hist"].items()}
    tpl = TPL[t]
    fx_line = (f"Reports in {L['ccy']}; all figures converted to USD at {L['fxr']:g} {L['ccy']}/USD (ECB, Oct 16, 2026), "
               f"history included.") if L["ccy"] != "USD" else None
    caveats = []
    if fx_line:
        caveats.append(fx_line)
    if "ADS ratio" in (L["shares_src"] or ""):
        caveats.append(f"Share count uses an {L['shares_src']}; it isn't in the XBRL cover page, so check it.")
    if years and years[-1] < 2025:
        caveats.append(f"Latest annual data in SEC XBRL is FY{years[-1]}.")
    if t == "QNCX":
        caveats.append("Renamed **IRulya Therapeutics (Nasdaq: IRLA)** on Oct 9, 2026 after a merger. The price is IRLA's close; "
                       "the share count and cash predate the merger, so the per-share value is stale.")
    vr = (m["C11"] / price) if isinstance(m.get("C11"), (int, float)) and price else None
    if K == "Net cash" and vr and vr > 2:
        caveats.append(f"**Check before trusting this:** the stock trades at {1 / vr:.0%} of its reported net cash. "
                       f"Possible causes: convertible preferred or pre-funded warrants not in the share count, a reverse "
                       f"split or offering after the cover date, or cash burned since {L.get('bs_end')}.")
    if K == "DCF" and vr and vr > 3:
        caveats.append("**Check before trusting this:** the value is more than 3× the price. The market may be pricing "
                       "events after the latest filings (reimbursement, guidance, financing) that reported history can't show.")
    if K == "rNPV" and vr and vr > 3:
        caveats.append("**Check before trusting this:** the rNPV is more than 3× the price. Small biotechs often trade far "
                       "below benchmark pipeline value: they need financing (dilution), and the benchmark can't see trial "
                       "quality, competition or a slow launch.")
    if L.get("bs_end") and L["bs_end"] < "2026-01-01":
        caveats.append(f"Latest balance sheet in SEC XBRL is {L['bs_end']}.")

    fm = {"Ticker": [t], "File Name": name, "CIK": cik, "SIC": sic, "Stock Exchange": [R["ex"]] if R.get("ex") else None,
          "Updated": TODAY, "Model": f"Stock Models/Pharma/{t}.xlsx"}
    B = []      # body lines (level-2 headings; demoted when appended to an existing note)
    if K == "DCF":
        S, npv = C["scen"], C["npv"]
        im = C["implied_m10"]
        fy10 = L["fyears"][-1]
        if isinstance(im, str):
            imp = (f"No year-{len(L['fyears'])} operating margin up to 100% reaches the price; the market is paying for "
                   f"growth or pipeline this model doesn't include." if im.startswith(">") else
                   f"Even a −50% margin by {fy10} values the stock above the price.")
        else:
            imp = (f"**The price implies a {fy10} operating margin of about {pct(im)}**, holding growth and the discount "
                   f"rate (base case {pct(P['m10'])}).")
        RV = REV.get(t) if P.get("review") else None
        if RV:
            o = C.get("opt", {})
            oi_ = {int(k): v for k, v in o.get("inputs", {}).items()}
            ob = (o.get("cases") or {}).get("Base", [None] * 4)
            B += [f"> [!important] Reviewed DCF (close of {TODAY})",
                  f"> - **{usd(m['C11'])} per share vs. {usd(price)} ({pct(m['C12'], sign=True)}).** Bear {usd(S['Bear']['ps'])}, "
                  f"bull {usd(S['Bull']['ps'])}. The basic screen said {usd(RV['basic']['ps'])} ({RV['basic']['up']}).",
                  f"> - **What changed:** revenue now follows guidance and patent expiries "
                  f"({', '.join(pct(g, 0, True) for g in RV['g'][:4])} for {L['fyears'][0]}–{L['fyears'][3]}), and the "
                  f"{fy10} margin target is {pct(P['m10'])} (basic: {pct(RV['basic']['m10'])}). Details below.",
                  f"> - {imp}",
                  f"> - **Options:** the market prices a ±{pct(oi_.get(20), 1)} move by {L['opt']['near']} and gives a "
                  f"{pct(ob[3])} chance of ending above the reviewed base value by {L['opt']['long']}."]
            for cv in caveats:
                B.append(f"> - {cv}")
            B.append("> - Model output, not a recommendation. I'm not a licensed advisor.")
            B.append("")
        B += [f"> [!important] Basic DCF (close of {TODAY})",
              (f"> - **{usd(m['C11'])} per share vs. {usd(price)} ({pct(m['C12'], sign=True)}).** " if m["C11"] >= 0 else
               f"> - **Below zero: {usd(m['C11'])} per share vs. {usd(price)}.** ")
              + 
              f"Bear {usd(S['Bear']['ps'])}, bull {usd(S['Bull']['ps'])}.",
              f"> - **Built from reported numbers only.** TTM revenue {money(L['rev_ttm'])} with a {pct(P['m_ttm'])} "
              f"operating margin. Revenue grows {pct(P['g1'], 1)} in the first year, fading to 2.5%; the margin moves "
              f"from {pct(P['m0'])} to {pct(P['m10'])} by {fy10}; discount rate {pct(P['disc'], 1)}.",
              f"> - {imp}",
              (f"> - **Terminal value is {pct(npv['tv_pct'])} of enterprise value.**" if npv["ev"] > 0 else
               "> - **Enterprise value is negative:** in the base case, the operating losses before the margin "
               "recovers outweigh the later profits.")
              + (" Equity is below zero after debt; in practice it would trade on option value."
                 if m["C11"] is not None and m["C11"] < 0 else "")]
        if P.get("m0_src", "TTM") != "TTM":
            B.append(f"> - The starting margin uses the {P['m0_src']}; TTM was {pct(P['m_ttm'])}.")
        for cv in caveats:
            B.append(f"> - {cv}")
        B.append("> - **Not in the model:** product-level forecasts, patent expiries, pipeline, guidance and deals. "
                 "Treat it as a screen, not a price target. I'm not a licensed advisor.")
        if RV:     # keep only the reviewed callout
            i0 = B.index(f"> [!important] Basic DCF (close of {TODAY})")
            j0 = next(j for j in range(i0, len(B)) if B[j] == "" or not B[j].startswith(">")) if any(
                B[j] == "" for j in range(i0, len(B))) else len(B)
            del B[i0:j0]
    elif K == "Net cash":
        why = (L["flag"] if L.get("flag") else
               "No revenue, so there are no cash flows to discount." if (L["rev_ttm"] or 0) < 5 else
               f"TTM revenue is {money(L['rev_ttm'], 1)}, below the \\$50M cutoff for a DCF.")
        evv = m["C9"]
        B += [f"> [!important] Net cash basis (close of {TODAY})",
              f"> - **{esc(why)}** The value shown is net cash: **{usd(m['C11'])} per share vs. {usd(price)} "
              f"({pct(m['C12'], sign=True)})**.",
              (f"> - **The market values the pipeline and business at {money(evv)}** (enterprise value)."
               if evv is not None and evv >= 0 else
               f"> - **The stock trades below net cash** (enterprise value {money(evv)}).")]
        cfo = m.get("C17")
        if isinstance(cfo, (int, float)):
            if cfo < 0:
                rw = m.get("C18")
                B.append(f"> - **Burn:** TTM operating cash flow {money(cfo, 1)}; net cash covers about "
                         f"{rw:.1f} years at that rate." if isinstance(rw, (int, float)) and rw > 0 else
                         f"> - **Burn:** TTM operating cash flow {money(cfo, 1)}, with net debt.")
            else:
                B.append(f"> - TTM operating cash flow is positive ({money(cfo, 1)}).")
        for cv in caveats:
            B.append(f"> - {cv}")
        B.append("> - A pipeline value needs program-level assumptions (peak sales, odds, timing), like the PDUFA "
                 "models. I'm not a licensed advisor.")
    elif K == "rNPV":
        S, n, PR = C["scen"], C["npv"], L["programs"]
        lead = PR[0]
        nxt = sorted((p["next_readout"], p["bucket"], p["stage"]) for p in PR if p.get("next_readout"))
        ip = C.get("implied_peak")
        imp = (f"**The price implies peak sales about {ip:.1f}× the indication benchmarks** (base case, BIO approval odds, "
               f"12% discount rate)." if isinstance(ip, (int, float)) else
               "**Even zero peak sales value the stock above its price** (it trades below the value of its cash after burn)."
               if ip == "below 0x" else "**The price implies peak sales more than 30× the indication benchmarks.**")
        B += [f"> [!important] Pipeline rNPV (close of {TODAY})",
              (f"> - **{usd(m['C11'])} per share vs. {usd(price)} ({pct(m['C12'], sign=True)}).** "
               if m["C11"] is not None and m["C11"] >= 0 else f"> - **Below zero: {usd(m['C11'])} per share vs. {usd(price)}.** ")
              + f"Bear {usd(S['Bear']['ps'])}, bull {usd(S['Bull']['ps'])}. Net cash alone is {usd(m.get('C13'))} per share.",
              f"> - **{len(PR)} program{'s' if len(PR) != 1 else ''} from ClinicalTrials.gov.** Lead: {lead['bucket']} "
              f"({STG.get(lead['stage'], lead['stage'])}, {pct(lead['pos'])} approval odds, launch ~{lead['launch']})."
              + (f" Next readout: ~{nxt[0][0]} ({nxt[0][1]})." if nxt else ""),
              f"> - {imp}"]
        cfo = m.get("C17")
        rw = m.get("C18")
        if isinstance(cfo, (int, float)) and cfo < 0:
            B.append(f"> - **Burn:** TTM operating cash flow {money(cfo, 1)}"
                     + (f"; net cash covers about {rw:.1f} years." if isinstance(rw, (int, float)) and rw > 0 else ", with net debt."))
        for p in PR:
            if p.get("reg"):
                B.append(f"> - **{p['bucket']}:** {esc(p['reg'])}.")
        for cv in caveats:
            B.append(f"> - {cv}")
        B.append("> - **Basic model:** indication-level peak sales and industry-average approval odds, not program-specific "
                 "forecasts. I'm not a licensed advisor.")
    else:
        miss = [x for x, v in (("share count", L["sh"]), ("cash", L["cash"])) if v is None]
        B += ["> [!important] No valuation",
              f"> - **SEC XBRL doesn't give a usable {' or '.join(miss)}**, so the workbook has history only."
              + (" " + esc(tpl["C4c"]) if tpl.get("C4c") and L["sh"] is None else "")]
        for cv in caveats:
            B.append(f"> - {cv}")
    B.append("")

    # snapshot
    B += ["## Snapshot", "", "| Item | Value | Note |", "| --- | --- | --- |",
          f"| Price | {usd(price)} | Close {TODAY} (Screener) |",
          f"| Shares | {L['sh']:,.1f}M | {esc(L['shares_src'] if L['shares_src'] != 'template' else 'SEC cover page')} |"
          if L["sh"] else "| Shares | n/a | not in SEC XBRL |",
          f"| Market cap | {money(m['C5'])} | |" if L["sh"] else "| Market cap | n/a | |",
          f"| Cash & investments | {money(L['cash'], 1)} | {srcl(L['cash_src'], L)} |",
          f"| Debt | {money(L['debt'], 1)} | {srcl(L['debt_src'], L)} |",
          f"| Enterprise value | {money(m['C9'])} | |" if L["sh"] and L["cash"] is not None else "| Enterprise value | n/a | |"]
    if years:
        B += [f"| TTM revenue | {money(L['rev_ttm'], 1)} | to {L['ttm_end']} |",
              f"| TTM operating income | {money(L['oi_ttm'], 1)} | |" if L.get("oi_ttm") is not None else None]
    B.append(f"| Balance sheet date | {L.get('bs_end') or 'n/a'} | |")
    B = [x for x in B if x is not None] + [""]

    if years:
        B += [f"## History (\\$M, FY{years[0]}–FY{years[-1]})", "",
              "| FY | Revenue | Growth | Operating income | Margin | Net income |", "| --- | --- | --- | --- | --- | --- |"]
        prev = None
        for y in years:
            h = hist[y]
            gr = pct(h["rev"] / prev - 1) if prev and prev > 0 and h["rev"] else ""
            mg = pct(h["oi"] / h["rev"]) if h["rev"] and h["rev"] > 0 else ""
            B.append(f"| {y} | {h['rev']:,.1f} | {gr} | {h['oi']:,.1f} | {mg} | {h['ni']:,.1f} |")
            prev = h["rev"]
        B.append("")

    if K == "DCF":
        S = C["scen"]
        sc = P["scen"]
        B += ["## Valuation", "", "| | Bear | Base | Bull |", "| --- | --- | --- | --- |"]
        for lab, key, dec in (("Discount rate", "disc", 1), ("Terminal growth", "tg", 1),
                              ("Revenue growth, first year", "g1", 1), (f"Operating margin, {L['fyears'][-1]}", "m10", 0)):
            B.append(f"| {lab} | " + " | ".join(pct(v, dec) for v in sc[key]) + " |")
        B.append("| Enterprise value | " + " | ".join(money(S[k]["ev"]) for k in ("Bear", "Base", "Bull")) + " |")
        B.append("| **NPV / share** | " + " | ".join(("**" + usd(S[k]["ps"]) + "**") if k == "Base" else usd(S[k]["ps"])
                                                       for k in ("Bear", "Base", "Bull")) + " |")
        B.append("| Upside | " + " | ".join(pct(S[k]["up"], sign=True) for k in ("Bear", "Base", "Bull")) + " |")
        B += ["", "Base-case forecast (\\$M):", "", "| FY | Revenue | Growth | Operating margin | Free cash flow | PV |",
              "| --- | --- | --- | --- | --- | --- |"]
        fc = C["fc"]
        for i, y in enumerate(L["fyears"]):
            B.append(f"| {y}E | {fc['rev'][i]:,.0f} | {pct(fc['g'][i], 1)} | {pct(fc['m'][i])} | {fc['fcf'][i]:,.0f} | "
                     f"{fc['pv'][i]:,.0f} |")
        n = C["npv"]
        B += ["", f"PV of forecast cash flow {money(n['pv_fcf'])} + PV of terminal value {money(n['pv_tv'])} = enterprise "
                  f"value {money(n['ev'])}; plus cash {money(L['cash'])}, less debt {money(L['debt'])} = equity "
                  f"{money(n['eq'])}.", "",
              "NPV / share by discount rate (rows) and terminal growth (columns):", "",
              "| | " + " | ".join(pct(v, 1) for v in L["tgs"]) + " |", "| --- |" + " --- |" * 5]
        for j, dr in enumerate(L["drs"]):
            B.append(f"| {pct(dr, 1)} | " + " | ".join(usd(v) for v in C["sens"][j]) + " |")
        B.append("")

    if K == "rNPV":
        PR, CP, S, n = L["programs"], C["programs"], C["scen"], C["npv"]
        B += ["## Pipeline", "",
              "One row per indication, from trials where the company is lead sponsor (or partner of an industry sponsor); "
              "terminated and withdrawn trials are excluded. Approval odds are the BIO 2011–2020 phase-transition rates for "
              "the disease area (rare and chronic diseases use the report's own rates).", "",
              "| Program (indication) | Drug(s) | Stage | Approval odds | Peak sales | Launch | Economics | Risked PV | Next readout | Trials |",
              "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |"]
        for p, c in zip(PR, CP):
            econ = ("owned, " + pct(p["share"]) + " net margin") if p["owned"] and p["share"] == 0.40 else \
                   ("partnered (" + ", ".join(p.get("partners", [])[:2]) + "), " + pct(p["share"])) if not p["owned"] else pct(p["share"])
            trials = " ".join(f"[{x}](https://clinicaltrials.gov/study/{x})" for x in p.get("lead_ncts", [])[:3])
            flags = (" (biomarker niche, 40% of bucket peak)" if p.get("niche") else "") + \
                    (f" (also {', '.join(p['label_expansions'])})" if p.get("label_expansions") else "")
            B.append(f"| {p['bucket']}{flags} | {esc(', '.join(p.get('drugs', [])[:2]))[:70]} | {STG.get(p['stage'], p['stage'])} | "
                     f"{pct(p['pos'])} | {money(p['peak'])} | {p['launch']} | {econ} | {money(c['pv'], 1)} | "
                     f"{p.get('next_readout') or ''} | {trials} |")
        meta = L.get("pipe_meta") or {}
        B += ["", f"Studies found: {meta.get('n_studies')} (sponsor search \"{meta.get('query')}\"); used: {meta.get('n_used')}."
              + (f" Programs beyond the top 8 not modeled: {', '.join(meta['dropped_programs'])}." if meta.get("dropped_programs") else ""), ""]
        B += ["## Valuation", "", "| | Bear | Base | Bull |", "| --- | --- | --- | --- |",
              "| Discount rate | 14% | 12% | 10% |", "| Peak-sales multiplier | 0.6× | 1.0× | 1.4× |",
              "| Approval-odds multiplier | 0.8× | 1.0× | 1.15× |",
              "| Pipeline rNPV (operations) | " + " | ".join(money(S[k]["ev"]) for k in ("Bear", "Base", "Bull")) + " |",
              "| **rNPV / share** | " + " | ".join(("**" + usd(S[k]["ps"]) + "**") if k == "Base" else usd(S[k]["ps"])
                                                  for k in ("Bear", "Base", "Bull")) + " |",
              "| Upside | " + " | ".join(pct(S[k]["up"], sign=True) for k in ("Bear", "Base", "Bull")) + " |", "",
              f"PV of risked pipeline contribution {money(n['pv_contrib'])} + PV of corporate burn {money(n['pv_burn'])} + "
              f"PV of taxes {money(n['pv_tax'])} = {money(n['ops'])}; plus cash {money(L['cash'])}, less debt "
              f"{money(L['debt'])} = equity {money(n['eq'])}.", "",
              "rNPV / share by approval-odds multiplier (rows) and peak-sales multiplier (columns):", "",
              "| | " + " | ".join(f"{v:g}×" for v in L["pipe"]["pk"]) + " |", "| --- |" + " --- |" * 5]
        for j, v in enumerate(L["pipe"]["ps"]):
            B.append(f"| {v:g}× | " + " | ".join(usd(x) for x in C["sens"][j]) + " |")
        B.append("")

    if K == "DCF" and P.get("review") and t in REV:
        RV = REV[t]
        B += [f"## Review ({TODAY})", "",
              "The basic screen extrapolated reported history. The review sets growth and margin from guidance, "
              "patent expiries and recent results:", ""]
        B += [f"- {w}" for w in RV["why"]]
        B += ["", "| | Basic screen | Reviewed bear | Reviewed base | Reviewed bull |", "| --- | --- | --- | --- | --- |",
              f"| NPV / share | {usd(RV['basic']['ps'])} | {usd(C['scen']['Bear']['ps'])} | **{usd(C['scen']['Base']['ps'])}** | "
              f"{usd(C['scen']['Bull']['ps'])} |",
              f"| {L['fyears'][-1]} operating margin | {pct(RV['basic']['m10'])} | " + " | ".join(pct(v) for v in RV["m10"]) + " |",
              "", "Reviewed revenue growth (base): " + " · ".join(f"{y} {pct(g, 1, True)}" for y, g in zip(L["fyears"], RV["g"])) + ".",
              "", "Review sources: " + " · ".join(f"[{a}]({u})" for a, u in RV["src"]), ""]
        o, lo = C["opt"], L["opt"]
        oi_ = {int(k): v for k, v in o["inputs"].items()}
        B += ["## Options", "",
              f"Cboe delayed quotes, close {lo['pdate']} (spot {usd(lo['spot'])}). Near expiry {lo['near']}, long expiry "
              f"{lo['long']}; risk-free 4.0%, dividend yield {pct(lo['q'], 1)}. Workbook sheet: Options.", "",
              "| Market read | Value |", "| --- | --- |",
              f"| ATM implied volatility, near / long | {pct(lo['ivn'], 1)} / {pct(lo['ivl'], 1)} ({lo['ivsrc'][list(lo['ivsrc'])[-1]]}) |",
              f"| Expected move by {lo['near']} (ATM straddle) | ±{pct(oi_[20], 1)} ({usd(oi_[21])}–{usd(oi_[22])}) |",
              f"| 1-SD move by {lo['long']} | ±{pct(oi_[23], 0)} |", "",
              f"| Case | Value / share | vs. spot | Annual return needed | Market P(above) by {lo['long']} |",
              "| --- | --- | --- | --- | --- |"]
        for n, v in o["cases"].items():
            B.append(f"| {n} | {usd(v[0])} | {pct(v[1], 0, True)} | {pct(v[2], 0, True) if isinstance(v[2], (int, float)) else ''} | {pct(v[3])} |")
        B += ["", f"Strategies at the {lo['long']} expiry (per contract; premium + = credit):", "",
              "| Strategy | Strikes | Premium | Max profit | Max loss | Breakeven | P/L at bear / base / bull | Market P(profit) |",
              "| --- | --- | --- | --- | --- | --- | --- | --- |"]
        lg = lo["legs"]
        strikes_ = {"Long ATM call": f"{lg['c0'][0]:g}C", "Bull call spread (ATM / upper)": f"{lg['c0'][0]:g}C / {lg['chi'][0]:g}C",
                    "Cash-secured put (~10% below spot)": f"{lg['pp'][0]:g}P",
                    "Covered call (100 shares + upper call, incl. dividends)": f"stock + {lg['chi'][0]:g}C",
                    "Bear put spread (ATM / lower)": f"{lg['p0'][0]:g}P / {lg['plo'][0]:g}P"}
        dl = lambda x: x if isinstance(x, str) else (("−" if x < 0 else "+" if x > 0 else "") + f"${abs(x):,.0f}")
        for lab, v in o["strat"].items():
            B.append(f"| {lab} | {strikes_[lab]} | {dl(v[0])} | {dl(v[1])} | {dl(v[2])} | {usd(v[3])} | "
                     f"{dl(v[4])} / {dl(v[6])} / {dl(v[7])} | {pct(v[8])} |")
        theo = [f"{lg[k][0]:g}{'C' if k[0] == 'c' else 'P'}" for k in lg if lg[k][2] == "theo"]
        B += ["", "- **Reading it:** probabilities are risk-neutral (lognormal at the long ATM IV), so they show what option "
              "prices imply, not a forecast. A low probability for a base value above spot means option prices put little "
              "weight on the model's outcome by that date.",
              f"- **Prices:** legs use the bid/ask mid when the spread is within 35% of mid; "
              + (f"Black-Scholes at the ATM IV for {', '.join(theo)} (wide quotes)." if theo else "all legs use market mids."),
              "- Model output, not a recommendation. I'm not a licensed advisor.", ""]
    B += ["## Model notes", ""]
    if K == "DCF":
        B += [f"- **Method:** free cash flow = operating income × (1 − 17% tax) − 25% of each year's revenue increase; "
              f"10 years, mid-year discounting from the balance sheet date (first year counts {P['stub']:.2f} of a year), "
              f"Gordon terminal value. Equity = EV + cash − debt.",
              f"- **Growth:** {esc(P['g_src'])}, capped at −10% to +35%, then a straight line to 2.5% terminal growth.",
              f"- **Margin:** starts at {pct(P['m0'])} and moves in a straight line to the target. The base target is "
              f"the median of TTM and the last three fiscal years ({pct(P['m_norm'])}), floored at {pct(P['floor'])} "
              f"(gross margin less 15 points, max 20%) and capped at 60%. Bear and bull are ±5 points.",
              f"- **Discount rate** by market cap ({money(L['mcap'])}): 8% above \\$100B, 8.5% \\$20–100B, 9.5% "
              f"\\$5–20B, 11% \\$1–5B, 12.5% below; ±1 point for bear/bull."]
    if K == "rNPV":
        B += ["- **Method:** each program's revenue = peak sales × launch ramp (10/25/45/65/85/100% over six years) × "
              "approval odds, with 12 years of exclusivity and 50% annual erosion after. Owned programs earn a 40% net "
              "margin (after COGS, SG&A and royalties); partnered programs 12% (royalty / profit share) unless noted.",
              f"- **Company costs:** the TTM operating loss ({money(L['burn'], 1)}) continues until the first launch "
              "(at most 4 years), then 30% of it (platform R&D and G&A), scaled by the chance at least one owned program "
              f"is approved. Taxes at 21% after a tax-loss carryforward of {money(L['nol'])}. Discount rate 12% (14% / 10% "
              "bear / bull), mid-year from the balance sheet date, 2026–2045.",
              "- **Approval odds:** BIO / Informa / QLS, *Clinical Development Success Rates 2011–2020* (Feb 2021): "
              "phase-transition rates by disease area, with Phase 1/2 and 2/3 interpolated, filed NDA/BLA at the area's "
              "approval rate, and approved products at 100%.",
              "- **Peak sales:** a table of average peaks for a successful launch in each indication (e.g. lung cancer "
              "$3B, obesity $6B, rare diseases $0.6–1B). Biomarker-selected oncology niches get 40%. These are "
              "placeholders; the Pipeline sheet's Peak column is the place to put program-specific estimates."]
    B += [f"- **History:** every fiscal year in SEC XBRL ({'FY' + str(years[0]) + '–FY' + str(years[-1]) if years else 'none'}). "
          f"\"SG&A & Other\" is the balancing line between gross profit, R&D and reported operating income (it "
          f"includes amortization, IPR&D and impairments). Net income is as reported." if years else
          "- **History:** no annual financials in SEC XBRL.",
          f"- **Workbook:** {C['n_cells']:,} cells recalculated with the Python `formulas` engine, 0 errors; "
          f"values cached so the Screener reads them without opening the file."]
    for fx_ in L.get("fixes", []):
        B.append(f"- **Data fix:** {fx_}.")
    for cell, lab in (("C4c", "Shares"), ("C7c", "Cash"), ("C8c", "Debt")):
        if tpl.get(cell):
            B.append(f"- **{lab} source:** {esc(tpl[cell])}")
    B += ["", "## Sources", ""]
    src = [f"[SEC XBRL company facts](https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json)"]
    src += [f"[{lab}]({u})" for lab, u in fl]
    if K == "rNPV":
        q = (L.get("pipe_meta") or {}).get("query") or name
        src.append(f"[ClinicalTrials.gov sponsor search](https://clinicaltrials.gov/search?spons={q.replace(' ', '%20')})")
        src.append("[BIO Clinical Development Success Rates 2011–2020](https://go.bio.org/rs/490-EHZ-999/images/ClinicalDevelopmentSuccessRates2011_2020.pdf)")
    src.append(f"[EDGAR filings](https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=&dateb=&owner=include&count=40)")
    B += ["- " + " · ".join(src), "", "Related: [[Basic NPV Models]] · [[Stock Model]]", ""]
    return fm, name, B


def yaml_val(v):
    if isinstance(v, list):
        return "\n" + "\n".join(f"  - {x}" for x in v)
    return f" {v}"


def write(t):
    fm, name, B = note(t)
    B = [re.sub(r"(?<!\\)\$", r"\\$", ln) for ln in B]
    path = f"{VAULT}/{t}.md"
    if t in USER_NOTES:        # notes the user wrote before the models: keep their text, replace the model section
        old = open(path).read()
        m = re.match(r"---\n(.*?)\n---\n?(.*)", old, re.S)
        head, body = (m.group(1), m.group(2)) if m else ("", old)
        keys = {}
        for ln in head.split("\n"):
            mm = re.match(r"([^\s:][^:]*):(.*)", ln)
            if mm:
                keys[mm.group(1).strip()] = mm.group(2).strip()
        lines = head.split("\n")
        for k, v in fm.items():
            if v is None:
                continue
            if k not in keys:
                lines.append(f"{k}:{yaml_val(v)}")
            elif k == "Ticker" and keys[k] == "" and not any(l.strip().startswith("- ") for l in
                                                              lines[lines.index("Ticker:") + 1:lines.index("Ticker:") + 2]):
                i = lines.index("Ticker:")
                lines[i] = f"Ticker:{yaml_val(v)}"
        body = re.sub(r"\n## Basic NPV model.*", "", body, flags=re.S).rstrip()
        sec = [("#" + ln) if ln.startswith("## ") else ln for ln in B]
        text = "---\n" + "\n".join(lines) + "\n---\n" + body + "\n\n## Basic NPV model (" + TODAY + ")\n\n" + "\n".join(sec)
    else:
        lines = [f"{k}:{yaml_val(v)}" for k, v in fm.items() if v is not None]
        text = "---\n" + "\n".join(lines) + "\n---\n" + f"# {name}\n\n" + "\n".join(B)
    text = re.sub(r"\n{3,}", "\n\n", text)
    open(f"{D}/notes/{t}.md", "w").write(text)


if __name__ == "__main__":
    os.makedirs(f"{D}/notes", exist_ok=True)
    for t in sys.argv[1:]:
        write(t)
