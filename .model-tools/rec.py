"""Long/short allocation of a 100% gross portfolio from the model values.  usage: rec.py universe.json OUT.json
Rules (also printed on the Recommendations sheet):
  universe   valuation models only (DCF, rNPV); net-cash and no-data rows are excluded
  liquidity  price >= $5 and market cap >= $1B (keeps micro-caps and penny stocks out)
  signal     long if upside >= +15%, short if upside <= -15%; upside clipped to +100% / -60%
  score      clipped upside x model confidence x size factor (confidence: reviewed/full DCF 1.0, rNPV 0.7,
             basic DCF 0.5, basic DCF with an operating loss today 0.25)
  books      60% long / 40% short of gross; weights within a book proportional to score
  caps       per-name cap by market cap (>= $50B 8%, $10-50B 6%, $3-10B 4%, $1-3B 2%), excess redistributed;
             positions under 0.5% dropped"""
import json, math, sys

CFG = {"min_price": 5.0, "min_mcap": 1000.0, "thresh": 0.15, "clip_hi": 1.0, "clip_lo": -0.6,
       "long_book": 0.60, "short_book": 0.40, "min_w": 0.005,
       "conf_lossmaking": 0.25, "conf": {"Reviewed DCF": 1.0, "Full DCF": 1.0, "PDUFA rNPV": 0.7, "rNPV": 0.7, "Basic DCF": 0.5},
       "caps": [(50000, 0.08), (10000, 0.06), (3000, 0.04), (1000, 0.02)]}


def size_factor(mcap):
    return 0.25 + 0.75 * min(1.0, max(0.0, (math.log10(mcap) - 3) / 2))


def cap_for(mcap):
    return next(c for lim, c in CFG["caps"] if mcap >= lim)


def allocate(names, book):
    """Water-fill: weights proportional to score, capped per name, excess redistributed; drop tiny weights."""
    live = dict(names)
    while True:
        w, fixed = {}, {}
        for _ in range(50):
            free = {t: s for t, s in live.items() if t not in fixed}
            rem = book - sum(fixed.values())
            tot = sum(s["score"] for s in free.values())
            if not free or tot <= 0 or rem <= 1e-12:
                break
            over = False
            for t, s in free.items():
                v = rem * s["score"] / tot
                if v > s["cap"] + 1e-12:
                    fixed[t] = s["cap"]; over = True
            if not over:
                for t, s in free.items():
                    w[t] = rem * s["score"] / tot
                break
        w.update(fixed)
        small = [t for t, v in w.items() if v < CFG["min_w"]]
        if not small:
            return w, book - sum(w.values())
        for t in small:
            live.pop(t)


def main(src, dst):
    U = json.load(open(src))
    elig, why = [], {}
    for r in U:
        if r["kind"] not in CFG["conf"] or r["value"] is None or r["up"] is None or r["mcap"] is None:
            continue
        if r["price"] < CFG["min_price"] or r["mcap"] < CFG["min_mcap"]:
            why[r["t"]] = "below $5 or under $1B"
            continue
        u = r["up"]
        if abs(u) < CFG["thresh"]:
            why[r["t"]] = "upside within ±15%"
            continue
        uc = min(u, CFG["clip_hi"]) if u > 0 else max(u, CFG["clip_lo"])
        sf, cf = size_factor(r["mcap"]), CFG["conf"][r["kind"]]
        if r["kind"] == "Basic DCF" and (r.get("m0") or 0) < 0:
            cf = CFG["conf_lossmaking"]    # margin path from losses to the 20% floor is the weakest part of a basic DCF
        elig.append(dict(r, uc=uc, sf=sf, cf=cf, score=abs(uc) * cf * sf, cap=cap_for(r["mcap"]),
                         side="Long" if u > 0 else "Short"))
    out = []
    cash = {}
    for side, book in (("Long", CFG["long_book"]), ("Short", CFG["short_book"])):
        names = {e["t"]: e for e in elig if e["side"] == side}
        w, left = allocate(names, book)
        cash[side] = left
        for t, v in w.items():
            out.append(dict(names[t], weight=v if side == "Long" else -v))
        for t in names:
            if t not in w:
                why[t] = "score too small (<0.5%)"
    out.sort(key=lambda e: (e["side"] != "Long", -abs(e["weight"])))
    json.dump({"cfg": CFG, "positions": out, "unallocated": cash, "excluded": why,
               "n_universe": sum(1 for r in U if r["kind"] in CFG["conf"])}, open(dst, "w"), indent=0, default=str)
    L = [e for e in out if e["side"] == "Long"]
    S = [e for e in out if e["side"] == "Short"]
    print(f"longs {len(L)} = {sum(e['weight'] for e in L):.1%} | shorts {len(S)} = {sum(e['weight'] for e in S):.1%} "
          f"| unallocated {cash}")
    for e in out:
        print(f"  {e['side']:5s} {e['t']:5s} {e['weight']:+.2%}  up {e['up']:+.0%}  {e['kind']:12s} mcap {e['mcap'] / 1000:7.1f}B "
              f"cap {e['cap']:.0%}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
