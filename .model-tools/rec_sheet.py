"""Write (or rewrite) the Recommendations sheet in the Screener from a rec.py result.
usage: rec_sheet.py SCREENER_IN OUT REC_JSON CLOSE_DATE(YYYY-MM-DD)"""
import json, sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

SRC, OUT, RECP, CLOSE = sys.argv[1:5]
REC = json.load(open(RECP))
wb = openpyxl.load_workbook(SRC)
ws = wb["Pharma"]
FONT = ws["B3"].font
keep_rows = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 2).value]
last = max(keep_rows)
keep = [{"vals": [ws.cell(r, c).value for c in range(2, 7)]} for r in keep_rows]
keep = [dict(x, t=str(x["vals"][0]).strip()) for x in keep]
# ---------------------------------------------------------------- Recommendations sheet
idx, keep_b3 = wb.sheetnames.index("Pharma") + 1, True
if "Recommendations" in wb.sheetnames:       # keep the user's tab position and their edits to the title rows
    idx = wb.sheetnames.index("Recommendations")
    keep_b3 = wb["Recommendations"]["B3"].value is not None
    del wb["Recommendations"]
R = wb.create_sheet("Recommendations", index=idx)
F = lambda bold=False, italic=False, color="000000", size=12: Font(name=FONT.name, size=size, bold=bold, italic=italic, color=color)
GREEN, RED, MUTED = "006100", "9C0006", "595959"
put = lambda a, v, fmt=None, **kw: (setattr(R[a], "value", v), setattr(R[a], "font", F(**kw)),
                                    fmt and setattr(R[a], "number_format", fmt))
R.column_dimensions["A"].width = 3
for col, w in zip("BCDEFGHIJKLMNOPQR", (5, 8, 34, 7, 9, 14, 12, 10, 9, 12, 11, 10, 9, 7, 11, 12, 60)):
    R.column_dimensions[col].width = w
put("B2", "Recommendations: model-weighted long/short portfolio (100% gross)", bold=True, size=14)
keep_b3 and put("B3", f"Weights from the Stock Models valuations at the {CLOSE} close. Model output, not investment advice; "
          "I'm not a licensed advisor.", italic=True, color=MUTED)
P = REC["positions"]
L = [p for p in P if p["side"] == "Long"]
S = [p for p in P if p["side"] == "Short"]
wl = sum(p["weight"] for p in L)
ws_ = -sum(p["weight"] for p in S)
summ = [("Long book", wl, "0.0%"), ("Short book", -ws_, "0.0%;-0.0%"), ("Gross exposure", wl + ws_, "0.0%"),
        ("Net exposure", wl - ws_, "0.0%"), ("Long positions", len(L), "0"), ("Short positions", len(S), "0"),
        ("Weighted upside, longs", sum(p["weight"] * p["up"] for p in L) / wl, "+0%;-0%"),
        ("Weighted upside, shorts", sum(-p["weight"] * p["up"] for p in S) / ws_, "+0%;-0%"),
        ("Valuation models in universe", REC["n_universe"], "0"),
        ("Positions after filters", len(P), "0")]
put("B5", "Summary", bold=True)
for i, (lab, v, fmt) in enumerate(summ):
    put(f"C{6 + i}", lab)
    R.merge_cells(f"C{6 + i}:E{6 + i}")
    put(f"F{6 + i}", v, fmt, bold=i < 4)
c = REC["cfg"]
rules = ["How the weights are set",
         "Universe: the 168 models with a valuation (DCF or rNPV). Net-cash and no-data tickers are left out.",
         f"Liquidity: price at least ${c['min_price']:.0f} and market cap at least ${c['min_mcap'] / 1000:.0f}B, so micro-caps and penny stocks can't drive the book.",
         f"Signal: long if upside is +{c['thresh']:.0%} or more, short if -{c['thresh']:.0%} or less; upside clipped to +100% / -60%.",
         "Score = upside × model confidence × size factor. Confidence: reviewed or full DCF 1.0, rNPV 0.7, basic DCF 0.5, basic DCF with an operating loss today 0.25.",
         "Size factor rises from 0.25 at $1B to 1.0 at $100B+ (log scale).",
         f"Books: {c['long_book']:.0%} long / {c['short_book']:.0%} short. Caps: 8% for $50B+, 6% for $10–50B, 4% for $3–10B, 2% for $1–3B; positions under 0.5% are dropped.",
         "Basic DCFs use GAAP margins and no pipeline, so they lean negative for acquisitive big pharma; that's why they get half weight."]
for i, s in enumerate(rules):
    put(f"H{5 + i}", s, bold=i == 0, color="000000" if i == 0 else MUTED)
h0 = 17
hd = ["#", "Ticker", "Company", "Side", "Weight", "Model", "Value / share", "Price", "Upside", "Mkt cap ($B)",
      "Confidence", "Size factor", "Cap", "Score", "Live upside", "Event", "Note(s)"]
thin = Border(bottom=Side(style="thin"))
for i, h in enumerate(hd):
    cell = R.cell(h0, 2 + i)
    cell.value = h
    cell.font = F(bold=True)
    cell.border = thin
    cell.alignment = Alignment(horizontal="left" if h in ("Ticker", "Company", "Model", "Note(s)") else "center")
names = {x["t"]: x["vals"][1] for x in keep}
r = h0 + 1
for grp, label in ((L, "Longs"), (S, "Shorts")):
    R.cell(r, 3).value = label
    R.cell(r, 3).font = F(bold=True)
    r += 1
    for j, p in enumerate(grp, start=1):
        vals = [j, p["t"], names.get(p["t"], ""), p["side"], p["weight"], p["kind"], p["value"], p["price"], p["up"],
                p["mcap"] / 1000, p["cf"], p["sf"], p["cap"], p["score"]]
        fmts = ["0", None, None, None, "0.0%;-0.0%", None, "$#,##0.00", "$#,##0.00", "+0%;-0%", "#,##0.0", "0.00",
                "0.00", "0%", "0.000"]
        for i, (v, fm) in enumerate(zip(vals, fmts)):
            cell = R.cell(r, 2 + i)
            cell.value = v
            cell.font = F(color=(GREEN if p["side"] == "Long" else RED) if i in (3, 4) else "000000", bold=i == 4)
            if fm:
                cell.number_format = fm
            cell.alignment = Alignment(horizontal="left" if i in (1, 2, 5) else "center")
        mt = f'MATCH($C{r},Pharma!$B$3:$B${last},0)'
        R.cell(r, 16).value = f'=IFERROR(INDEX(Pharma!$K$3:$K${last},{mt}),"")'
        R.cell(r, 16).number_format = "+0%;-0%"
        R.cell(r, 17).value = f'=IFERROR(IF(INDEX(Pharma!$M$3:$M${last},{mt})="","",INDEX(Pharma!$M$3:$M${last},{mt})),"")'
        R.cell(r, 17).number_format = "yyyy-mm-dd"
        R.cell(r, 18).value = f'=IFERROR(T(INDEX(Pharma!$N$3:$N${last},{mt})),"")'
        for cc in (16, 17, 18):
            R.cell(r, cc).font = F()
            R.cell(r, cc).alignment = Alignment(horizontal="left" if cc == 18 else "center")
        r += 1
    R.cell(r, 4).value = f"Total {label.lower()}"
    R.cell(r, 4).font = F(bold=True)
    R.cell(r, 6).value = sum(p["weight"] for p in grp)
    R.cell(r, 6).number_format = "0.0%;-0.0%"
    R.cell(r, 6).font = F(bold=True, color=GREEN if label == "Longs" else RED)
    R.cell(r, 6).border = Border(top=Side(style="thin"))
    R.cell(r, 6).alignment = Alignment(horizontal="center")
    r += 2
R.freeze_panes = f"B{h0 + 1}"
R.sheet_view.showGridLines = False
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print(f"Recommendations: {len(L)} long / {len(S)} short, as of {CLOSE}")
