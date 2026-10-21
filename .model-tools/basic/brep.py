"""Large-cap upside review: chart + markdown report.  usage: brep.py IMAGE_NUMBER -> report/ (md + png)"""
import json, os, sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

D = os.path.dirname(os.path.abspath(__file__))
NUM = int(sys.argv[1])
# close date: --close YYYY-MM-DD, else the close the prices in screener_cache.json are from
CLOSE = sys.argv[sys.argv.index("--close") + 1] if "--close" in sys.argv else \
    json.load(open(f"{D}/screener_cache.json"))["close"]
IMG = f"Image-{NUM:05d}"
os.makedirs(f"{D}/report", exist_ok=True)
REV = json.load(open(f"{D}/review.json"))
ROWS = {x["t"]: x for x in json.load(open(f"{D}/tickers.json"))}
TICK = list(REV)
DRIVER = {
    "NVO": "2026 guidance (0% to −6%) and the 2027 US semaglutide price resets",
    "PFE": "about \\$17B of revenue losing exclusivity through 2030",
    "GSK": "dolutegravir's 2028–2030 patent expiry, offset by guidance growth and a margin GSK says holds",
    "SNY": "raised guidance (about +10%) and Dupixent growth to 2030, before its US expiry around 2031",
    "ZTS": "the August guidance cut (organic −3% to −1%)",
    "INCY": "Jakafi's US patent expiry in December 2028",
    "HALO": "ENHANZE royalty step-downs as the rHuPH20 patents expire (2027–2029)",
    "NBIX": "Crenessity and VYKAT growth, with a possible Ingrezza Medicare price from 2029",
    "RDY": "generic Revlimid's fall-off and stale XBRL data (FY2026 actual +3.2%, Q1 FY2027 −5.6%)"}

D_ = {}
for t in TICK:
    L, C = json.load(open(f"{D}/out/{t}_layout.json")), json.load(open(f"{D}/out/{t}_calc.json"))
    o = C["opt"]
    oi = {int(k): v for k, v in o["inputs"].items()}
    D_[t] = dict(name=ROWS[t]["name"], price=L["price"], mcap=L["mcap"], basic=REV[t]["basic"]["ps"],
                 bear=C["scen"]["Bear"]["ps"], base=C["scen"]["Base"]["ps"], bull=C["scen"]["Bull"]["ps"],
                 up=C["main"]["C12"], im=C["implied_m10"], m10=L["P"]["m10"], move=oi[20], near=L["opt"]["near"],
                 long=L["opt"]["long"], pbase=o["cases"]["Base"][3], ivl=L["opt"]["ivl"], ivn=L["opt"]["ivn"],
                 spot=L["opt"]["spot"], strat=o["strat"], legs=L["opt"]["legs"], q=L["opt"]["q"], fy10=L["fyears"][-1])
order = sorted(TICK, key=lambda t: -D_[t]["up"])
OPT_DATE = max(json.load(open(f"{D}/out/{t}_layout.json"))["opt"]["pdate"] for t in TICK)   # option chains aren't refreshed daily

# ------------------------------------------------------------------ chart
SURF, INK, INK2, MUTED, GRID, AXIS = "#fcfcfb", "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
BLUE, ORANGE, GREEN = "#2a78d6", "#eb6834", "#1baf7a"
try:
    font_manager.fontManager.addfont("/System/Library/Fonts/HelveticaNeue.ttc")
except Exception:
    pass
plt.rcParams.update({"font.family": "Helvetica Neue", "font.size": 10.5, "text.color": INK, "axes.edgecolor": AXIS,
                     "axes.labelcolor": INK2, "xtick.color": MUTED, "ytick.color": INK2, "axes.facecolor": SURF,
                     "figure.facecolor": SURF, "savefig.facecolor": SURF, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.spines.left": False, "xtick.major.size": 0, "ytick.major.size": 0})
fig = plt.figure(figsize=(11.5, 5.6), dpi=200)
fig.text(0.03, 0.965, "Large-cap upside review: what survives a guidance check", fontsize=14, fontweight="bold", va="top")
fig.text(0.03, 0.91, "Left: value vs. price (0% = price), reviewed bear–bull range, reviewed base and the basic screen. "
         "Right: option-implied chance of ending above the reviewed base by the long expiry.", fontsize=9.5, color=INK2, va="top")
ax = fig.add_axes([0.08, 0.1, 0.58, 0.72])
lo, hi = -1.0, 2.0
for i, t in enumerate(reversed(order)):
    d = D_[t]
    b0, b1 = d["bear"] / d["price"] - 1, d["bull"] / d["price"] - 1
    ax.plot([max(b0, lo), min(b1, hi)], [i, i], color=BLUE, alpha=0.22, lw=9, solid_capstyle="round", zorder=1)
    if b1 > hi:
        ax.annotate(f"bull {b1:+.0%}", (hi, i), xytext=(-2, 7), textcoords="offset points", ha="right", fontsize=8, color=MUTED)
    bv = d["basic"] / d["price"] - 1
    ax.scatter([min(bv, hi)], [i], s=70, facecolor=SURF, edgecolor=ORANGE, linewidth=2, zorder=3)
    ax.scatter([d["up"]], [i], s=70, color=BLUE, edgecolor=SURF, linewidth=2, zorder=4)
    ax.annotate(f"{d['up']:+.0%}", (d["up"], i), xytext=(0, -13), textcoords="offset points", ha="center",
                fontsize=8.5, color=INK, fontweight="bold")
ax.axvline(0, color=INK2, lw=1, zorder=2)
ax.set_yticks(range(len(order)))
ax.set_yticklabels(list(reversed(order)))
ax.set_xlim(lo - 0.05, hi + 0.05)
ax.set_ylim(-0.7, len(order) - 0.3)
ax.set_xticks([-1, -0.5, 0, 0.5, 1, 1.5, 2])
ax.set_xticklabels(["−100%", "−50%", "price", "+50%", "+100%", "+150%", "+200%"])
ax.grid(axis="x", color=GRID, lw=0.8)
ax.set_axisbelow(True)
h1 = ax.scatter([], [], s=70, color=BLUE, label="Reviewed base")
h2 = ax.scatter([], [], s=70, facecolor=SURF, edgecolor=ORANGE, linewidth=2, label="Basic screen")
h3 = ax.plot([], [], color=BLUE, alpha=0.22, lw=9, label="Reviewed bear–bull")[0]
ax.legend(handles=[h1, h2, h3], loc="lower right", frameon=False, fontsize=9)
ax2 = fig.add_axes([0.74, 0.1, 0.23, 0.72])
for i, t in enumerate(reversed(order)):
    p = D_[t]["pbase"]
    ax2.barh(i, p, color=GREEN, height=0.55)
    ax2.text(p + 0.02, i, f"{p:.0%}", va="center", fontsize=9, color=INK)
ax2.set_yticks(range(len(order)))
ax2.set_yticklabels([])
ax2.set_xlim(0, 1.1)
ax2.set_ylim(-0.7, len(order) - 0.3)
ax2.set_xticks([0, 0.5, 1])
ax2.set_xticklabels(["0%", "50%", "100%"])
ax2.grid(axis="x", color=GRID, lw=0.8)
ax2.set_axisbelow(True)
ax2.set_title("P(above reviewed base)", fontsize=10, color=INK2, loc="left")
fig.savefig(f"{D}/report/{IMG}.png")

# ------------------------------------------------------------------ report
usd = lambda v: ("−" if v < 0 else "") + f"\\${abs(v):,.2f}"
pct = lambda v, s=True: ("−" if v < 0 else "+" if s and v > 0 else "") + f"{abs(v) * 100:.0f}%"
dl = lambda x: x if isinstance(x, str) else (("−" if x < 0 else "+" if x > 0 else "") + f"\\${abs(x):,.0f}")
survive = [t for t in order if D_[t]["up"] > 0.10]
gone = [t for t in order if D_[t]["up"] <= 0.10]
lines = ["---", f"Updated: {CLOSE}", "---", "# Large-Cap Upside Review", "",
         "The [[Basic NPV Models]] screen built DCFs from reported numbers only. This review takes the nine names with a "
         "market cap of at least \\$10B and at least 25% upside in that screen, checks each against current guidance, patent "
         "expiries and recent results, re-runs the DCF with reviewed assumptions, and adds an options model (Options sheet in "
         "each workbook). " + (f"Prices and option quotes are the {CLOSE} close." if CLOSE == OPT_DATE else
                               f"Prices are the {CLOSE} close; option quotes are the {OPT_DATE} close."), "",
         "> [!important] Findings",
         f"> - **The upside holds for {len(survive)} of 9: {', '.join(survive)}.** "
         + "; ".join(f"{t} {usd(D_[t]['base'])} ({pct(D_[t]['up'])})" for t in survive) + ".",
         f"> - **It disappears or reverses for {len(gone)}: {', '.join(gone)}.** The basic screen extrapolated growth that "
         "guidance or a patent cliff has already ruled out. NVO goes from \\$77 to \\$37 (at the price), INCY from \\$200 to "
         "\\$89, HALO from \\$183 to \\$65.",
         "> - **Options put little weight on the surviving upside.** The option-implied chance of ending above the reviewed "
         "base by the long expiry is " + ", ".join(f"{t} {D_[t]['pbase']:.0%}" for t in survive) +
         ". For names where the reviewed value is below the price, the chance of staying above it is " +
         ", ".join(f"{t} {D_[t]['pbase']:.0%}" for t in gone) + ".",
         "> - **These are model outputs, not recommendations.** I'm not a licensed advisor.", "",
         f"![[{IMG}.png|center]]", "",
         "## Summary", "",
         "| Ticker | Company | Price | Basic screen | Reviewed bear / base / bull | Reviewed upside | Price-implied "
         "margin | Expected move (near) | P(above base), long expiry |",
         "| --- | --- | --- | --- | --- | --- | --- | --- | --- |"]
for t in order:
    d = D_[t]
    im = d["im"] if isinstance(d["im"], str) else f"{d['im']:.0%} by {d['fy10']}"
    lines.append(f"| [[{t}]] | {d['name']} | {usd(d['price'])} | {usd(d['basic'])} | {usd(d['bear'])} / **{usd(d['base'])}** / "
                 f"{usd(d['bull'])} | {pct(d['up'])} | {im} | ±{d['move'] * 100:.1f}% to {d['near']} | "
                 f"{d['pbase']:.0%} by {d['long']} |")
lines += ["", "Price-implied margin: the year-10 operating margin that makes the reviewed DCF equal the price, holding the "
          "reviewed growth path and discount rate.", "", "## Company by company", ""]
for t in order:
    d, R = D_[t], REV[t]
    verdict = ("the upside survives" if d["up"] > 0.10 else "the upside is gone" if d["up"] > -0.10 else "the reviewed value is below the price")
    lines += [f"### [[{t}]] — {d['name']}", "",
              f"- **{usd(d['basic'])} → {usd(d['base'])} ({pct(d['up'])} vs. {usd(d['price'])}); {verdict}.** Main driver: "
              f"{DRIVER[t]}. Bear {usd(d['bear'])}, bull {usd(d['bull'])}.",
              f"- {R['why'][0]}"]
    if len(R["why"]) > 1:
        lines.append(f"- {R['why'][1]}")
    s = d["strat"]
    lg = d["legs"]
    lines += [f"- **Options ({d['long']}):** ATM IV {d['ivl']:.0%}; a {lg['c0'][0]:g}/{lg['chi'][0]:g} bull call spread costs "
              f"{dl(s['Bull call spread (ATM / upper)'][0])} for up to {dl(s['Bull call spread (ATM / upper)'][1])} "
              f"(market P(profit) {s['Bull call spread (ATM / upper)'][8]:.0%}); a {lg['pp'][0]:g} cash-secured put collects "
              f"{dl(s['Cash-secured put (~10% below spot)'][0])} (breakeven {usd(s['Cash-secured put (~10% below spot)'][3])}); a "
              f"{lg['p0'][0]:g}/{lg['plo'][0]:g} bear put spread costs {dl(s['Bear put spread (ATM / lower)'][0])}. "
              f"Full table in the note.", ""]
lines += ["## Options summary", "",
          "Payoffs at expiry per contract. \"At base\" assumes the stock reaches the reviewed base value by the long expiry, "
          "which is a strong assumption for a fundamental value.", "",
          "| Ticker | Long expiry | Bull call spread: cost / max / at base | Cash-secured put: credit / breakeven | Covered call: "
          "at base | Bear put spread: cost / at bear |", "| --- | --- | --- | --- | --- | --- |"]
for t in order:
    s, d = D_[t]["strat"], D_[t]
    b, c, cc, bp = (s["Bull call spread (ATM / upper)"], s["Cash-secured put (~10% below spot)"],
                    s["Covered call (100 shares + upper call, incl. dividends)"], s["Bear put spread (ATM / lower)"])
    lines.append(f"| {t} | {d['long']} | {dl(b[0])} / {dl(b[1])} / {dl(b[6])} | {dl(c[0])} / {usd(c[3])} | {dl(cc[6])} | "
                 f"{dl(bp[0])} / {dl(bp[4])} |")
lines += ["", "## Method", "",
          "- **Selection:** basic-screen DCFs with market cap ≥ \\$10B and upside ≥ 25%. Just below the cut: ALNY (+24%), "
          "BMY (+20%), NVS (+18%), ONC (+17%).",
          "- **Review:** each workbook keeps the basic structure (FCF = operating income × (1 − 17%) − 25% of revenue growth, "
          "10 years, Gordon terminal value, discount 8–9.5% by size). The base case gets a year-by-year revenue path from "
          "guidance and patent dates (Assumptions row 46), plus a reviewed margin start and target. Bear and bull shift "
          "first-year growth by ±5 points, fading to zero by year 10, and the margin target by ±5 points.",
          "- **GAAP basis:** margins are GAAP operating margins, so amortization of acquired intangibles counts as a cost. "
          "That's a conservative stand-in for the business development spending that refills pipelines. Company \"core\" or "
          "\"business\" margins are 5–10 points higher.",
          f"- **Options:** Cboe delayed quotes ({OPT_DATE} close). Legs use the mid when the spread is within 35% of mid, "
          "otherwise Black-Scholes-Merton at the ATM IV with the dividend yield. Probabilities are risk-neutral (lognormal at "
          "the long ATM IV), so they describe option prices, not real-world odds. RDY's chain is too thin for quotes; its long "
          "IV is 6-month realized volatility.",
          "- **Verification:** each workbook recalculated with the Python `formulas` engine, 0 errors, values cached.", "",
          "## Sources", ""]
for t in order:
    lines.append(f"- **{t}:** " + " · ".join(f"[{a}]({u})" for a, u in REV[t]["src"]))
lines += ["- Option quotes: [Cboe delayed quotes](https://www.cboe.com/delayed_quotes/)", "",
          "Related: [[Basic NPV Models]] · [[PDUFA Calendar]] · [[Black–Scholes Model]]", ""]
open(f"{D}/report/Large-Cap Upside Review.md", "w").write("\n".join(lines))
print("ok", IMG, order)
