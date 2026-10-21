"""Synthetic multi-day test of paper_trader.py in a temp folder (no network, no real ledger)."""
import datetime as dt, json, os, subprocess, sys, tempfile
import openpyxl

PY, HERE = sys.executable, os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp()
ENV = dict(os.environ, TRADER_DIR=TMP)
today = dt.date.today()


def sheet(targets, asof):
    wb = openpyxl.Workbook()
    wb.active.title = "Pharma"
    R = wb.create_sheet("Recommendations")
    R["B3"] = f"Weights from the Stock Models valuations at the {asof.isoformat()} close. Model output."
    R["C17"], R["E17"], R["F17"] = "Ticker", "Side", "Weight"
    r = 19
    for t, w in targets.items():
        R.cell(r, 3).value, R.cell(r, 5).value, R.cell(r, 6).value = t, "Long" if w > 0 else "Short", w
        r += 1
    p = os.path.join(TMP, "s.xlsx")
    wb.save(p)
    return p


def prices(d):
    p = os.path.join(TMP, "px.json")
    json.dump(d, open(p, "w"))
    return p


def run(*args, ok=True):
    r = subprocess.run([PY, os.path.join(HERE, "paper_trader.py"), *args], env=ENV, capture_output=True, text=True)
    assert (r.returncode == 0) == ok, (args, r.stdout, r.stderr)
    return r.stdout + r.stderr


def ledger():
    return json.load(open(os.path.join(TMP, "ledger.json")))


def set_last_run(days_ago):
    L = ledger()
    L["last_run"] = (today - dt.timedelta(days=days_ago)).isoformat()
    json.dump(L, open(os.path.join(TMP, "ledger.json"), "w"))


s5 = 5e-4
# day 1: open the book
out = run("run", "--sheet", sheet({"A": 0.3, "B": 0.3, "C": -0.2, "D": -0.2}, today), "--prices",
          prices({"A": 100, "B": 50, "C": 20, "D": 10}))
L = ledger()
assert L["positions"] == {"A": 300, "B": 600, "C": -1000, "D": -2000}, L["positions"]
slip = (30000 + 30000 + 20000 + 20000) * s5
assert abs(L["cash"] - (100000 - 60000 + 40000 - slip)) < 1e-6, L["cash"]
nav1 = L["cash"] + 300 * 100 + 600 * 50 - 1000 * 20 - 2000 * 10
assert abs(nav1 - (100000 - slip)) < 1e-6
print("day 1 ok: 4 opens, NAV", round(nav1, 2))
# dry run changes nothing; running twice the same day is refused
before = json.dumps(ledger())
run("run", "--dry-run", "--sheet", sheet({"A": 0.5, "C": -0.5}, today), "--prices", prices({"A": 100, "B": 50, "C": 20, "D": 10}))
assert json.dumps(ledger()) == before
run("run", "--sheet", sheet({"A": 0.3}, today), "--prices", prices({"A": 100, "B": 50, "C": 20, "D": 10}), ok=False)
print("dry run and once-a-day guard ok")
# day 2 (two calendar days later): B exits, E opens, D flips short -> long, A within the band stays, borrow accrues
set_last_run(2)
L0 = ledger()
px2 = {"A": 101, "B": 50, "C": 22, "D": 10, "E": 25}
out = run("run", "--sheet", sheet({"A": 0.3, "C": -0.2, "D": 0.2, "E": 0.3}, today), "--prices", prices(px2))
L = ledger()
fee = (1000 * 22 + 2000 * 10) * 0.02 * 2 / 360
nav_pre = L0["cash"] - fee + 300 * 101 + 600 * 50 - 1000 * 22 - 2000 * 10
assert "B" not in L["positions"], L["positions"]
def expect(cur, w, p):
    tgt = int(w * nav_pre / p)
    return cur if abs(cur - tgt) * p / nav_pre < 0.0075 or abs(cur - tgt) * p < 250 else tgt
assert L["positions"]["A"] == expect(300, 0.3, 101), (L["positions"]["A"], int(0.3 * nav_pre / 101))
assert L["positions"]["D"] == int(0.2 * nav_pre / 10), (L["positions"]["D"], nav_pre)
assert L["positions"]["E"] == int(0.3 * nav_pre / 25)
assert L["positions"]["C"] == expect(-1000, -0.2, 22), L["positions"]["C"]
acts = [(t["action"], t["ticker"]) for t in L["trades"] if t["date"] == today.isoformat()][-6:]
assert ("BUY TO COVER", "D") in acts and ("BUY", "D") in acts and ("SELL", "B") in acts and ("BUY", "E") in acts, acts
assert abs(L["fees_paid"] - fee) < 1e-6
print("day 2 ok: exit, new, flip, band, borrow fee", round(fee, 2))
# day 3: D moves +600% (looks like a split) -> held back; the split command fixes the share count
set_last_run(1)
d_before = ledger()["positions"]["D"]
out = run("run", "--sheet", sheet({"A": 0.3, "C": -0.2, "D": 0.2, "E": 0.3}, today), "--prices",
          prices({"A": 101, "C": 22, "D": 70, "E": 25}))
assert "Held back D" in out and ledger()["positions"]["D"] == d_before
row = open(os.path.join(TMP, "nav.csv")).read().strip().splitlines()[-1].split(",")
assert 0.95 <= float(row[5].rstrip("%")) / 100 <= 1.05 and float(row[2]) > 0, row
run("split", "D", "0.142857")
assert ledger()["positions"]["D"] == int(d_before * 0.142857)
print("split guard ok")
# stale sheet is refused, accepted with --force
set_last_run(1)
run("run", "--sheet", sheet({"A": 0.3}, today - dt.timedelta(days=9)), "--prices", prices({"A": 101, "C": 22, "D": 70, "E": 25}),
    ok=False)
print("stale-sheet guard ok")
for f in ("orders", "positions.csv", "nav.csv"):
    assert os.path.exists(os.path.join(TMP, f)), f
print(open(os.path.join(TMP, "nav.csv")).read())
print("ALL TESTS PASSED")
