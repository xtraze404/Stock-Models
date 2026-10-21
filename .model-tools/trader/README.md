# Paper trader

Rebalances a simulated account to the Screener's **Recommendations** sheet once a day.

Schwab's Trader API only connects to live accounts. It can't place orders in thinkorswim paperMoney, and its sandbox doesn't simulate fills. So the trader keeps its own ledger, starting at $100,000 like paperMoney, and writes each day's orders so you can copy them into paperMoney by hand if you want both to match.

## Daily use

After the Screener refresh, which updates prices and the Recommendations sheet:

```bash
cd ~/Documents/"Stock Models"/.model-tools/trader && ../venv/bin/python paper_trader.py run
```

- `--dry-run` shows the orders without saving anything.
- `--prices FILE.json` fills at a saved price file (for example the refresh's `closes_<date>.json`) instead of fetching live quotes.
- `status` lists positions and value at the last run's prices.
- `split TICKER RATIO` records a split (new shares per old share: 1-for-7 reverse = `0.142857`).
- `reset` starts over; the old ledger is kept as a `.bak.json`.

## What a run does

1. Reads the target weights from the Recommendations sheet (Longs positive, Shorts negative) and the price date from the Pharma tab's "Close" header.
2. Gets a price for every target and held name: the Nasdaq close after 4:00 pm ET, otherwise the latest quote, with Cboe as fallback.
3. Charges short borrow fees for the days since the last run (2% a year by default).
4. Plans whole-share orders:
   - Names no longer on the sheet are closed.
   - New names are opened, and a long that becomes a short (or the reverse) is closed and reopened.
   - Other positions are resized only when they're more than 0.75% of the account away from target and the trade is over $250. This keeps daily churn down.
5. Fills at the price ± 5 bp of slippage, then updates `ledger.json`, `positions.csv` and `nav.csv`, and writes `orders/<date>.csv` (BUY / SELL / SELL SHORT / BUY TO COVER).

## Safety checks

- Won't trade twice on the same day, or on a sheet more than 4 days old, unless you pass `--force`.
- Stops if more than 10% of tickers have no price. A held name with no price keeps its last price and isn't traded.
- A held name that moves more than 60% since the last run is held back as a possible split, and valued at its last price until you record the split with `split`.

All settings are in `config.json`, including the band, minimum trade size, slippage and borrow rate (with per-ticker overrides for hard-to-borrow names).

## Not modeled

Dividends (paid on longs, owed on shorts), margin interest and hard-to-borrow recalls. Orders fill at the run's price, not the next open.

`test_paper_trader.py` runs a multi-day simulation in a temp folder covering fills, exits, flips, the drift band, borrow fees and the guards.

Model output, not investment advice.
