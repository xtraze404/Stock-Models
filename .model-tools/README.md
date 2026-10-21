# Model tools

The Python pipeline behind `Pharma/*.xlsx`, the vault notes and `Screener.xlsx`. Run everything with the venv's Python, `.model-tools/venv/bin/python`.

## Daily refresh

After 4:00 pm ET on a trading day:

```bash
cd ~/Documents/"Stock Models"/.model-tools && venv/bin/python refresh_daily.py
```

One command runs the whole post-close update and ends with the paper trader. The rebuild takes about 5 minutes. Fetching closes adds 4 to 30 minutes, depending on how hard Nasdaq throttles (a re-fetch at 11:30 pm ET took 28 minutes). Close `Screener.xlsx` in Excel first. If it's open, the Screener, Recommendations and trader steps are skipped, and you finish with `--from screener` once it's closed.

| Option | What it does |
| --- | --- |
| `--dry-run` | Prints every step and its checks, and writes nothing. |
| `--date YYYY-MM-DD` | Uses a past close, from the historical feed. The default is today, and it refuses to run before 4:00 pm ET or on weekends. |
| `--accept-changes` | Installs even if a model's NPV/share changed (see step 4). |
| `--from STEP`, `--to STEP` | Runs part of the pipeline. Step names: `closes splits prices validate notes install screener recs trade`. |
| `--refetch` | Fetches all closes again, even if `closes_<yyyymmdd>.json` already exists. |
| `--no-trade`, `--force-trade` | Skips the paper trader, or passes `--force` to it. |
| `--jobs N` | Sets how many builds run in parallel (default 6). |

### Steps

1. **Closes.** Runs `closes_live.py "<Mon D, YYYY>" <iso> closes_<yyyymmdd>.json <tickers>` for every ticker in `basic/tickers.json`, using the quote feed's "Closed at" price. Nasdaq's historical feed posts official closes only late at night. Missing tickers are re-fetched once at 3 threads, then tried on the historical feed (`closes.py`). QNCX is priced as IRLA.
   - Stops if more than 5% of tickers have no close (a holiday, or the feed isn't posting yet).
   - Stops if most prices equal the previous close (a stale feed), or if the date is older than the loaded close.
   - Tickers that are still missing keep their previous price.
   - The quote feed's "Closed at" price can change later that evening. On 2026-10-20, 59 of 703 differed between 8:55 pm and 11:30 pm ET (median 0.6%, up to ±10% on sub-\$1 names). To use official closes, run the next morning with `--date <that day> --refetch --no-trade`, which uses the historical feed and leaves the trader alone.
2. **Split check.** Every move over 60% is checked against Nasdaq's quote JSON: a "Stock Split : <date>" notice dated after the previous close, plus the split-adjusted PreviousClose, which gives the ratio. Confirmed splits get `basic/templates.json` C4 adjusted the way REVB (1-for-7, `=(3.983416)/7`) and MBRX (1-for-20) were on 2026-10-20, and each adjustment is printed.
   - Applied splits are recorded in `basic/splits_applied.json`, so a re-run never adjusts twice.
   - A split in a hand-built model, or one with an unclear ratio, stops the run so you can fix it by hand.
   - A big move without a split notice is reported and used as is.
3. **Prices and rebuild.** Writes the closes into `basic/tickers.json` and `basic/screener_cache.json`. The cache keeps its row numbers, and its `"close"` field records the date. Then it runs `bbuild.py` + `bcalc.py` for every ticker in `basic/todo.txt`, 6 at a time, and `reprice.py` for the 16 hand-built models (LLY, EVMN and the 14 PDUFA names), repricing each installed workbook in place.
4. **Validation.** Compares each model's NPV/share (Main C11) with the installed workbook's. A price update shouldn't move it: discount rates are pinned in `basic/disc_pins.json`. If any value changed, other than split-adjusted names, the run stops before installing anything. Check the list, then run `--from validate --accept-changes` if the changes are expected.
5. **Notes.** Runs `bnote.py`, `bhub.py` and `brep.py` with `--close <date>`. Without `--close`, these scripts use the `"close"` in `screener_cache.json`. The Large-Cap report keeps its image number (Image-00937). Its option quotes keep the date of the stored chains, and its narrative text is fixed.
6. **Install.** Copies workbooks to `Pharma/`, notes to `Vault/Finance/Companies/`, and the hub, the report and its image to `Vault/Finance/` and `Gallery/`. It skips:
   - any workbook with a `~$` Excel lock file;
   - any file edited since the last install. `install_manifest.json` holds each file's MD5 at install, so your edits stay and the file is listed as skipped;
   - hand-built models and your own notes (MRK, LENZ, MNKD, SPRB), only if they changed after this run read them. Your edits to those carry forward.
7. **Screener.** Skipped if `~$Screener.xlsx` exists. Otherwise it backs up the Screener to `screener/backups/`, sets the Pharma prices in column F (matched by ticker, so re-sorting is fine) and the F1 header "Close YYYY-MM-DD", then runs `bscreener.py` (layout v2: J/K/L) and `fixlinks_all.py`.
   - It first checks that the columns are still where layout v2 puts them.
   - It installs only if a cell-by-cell comparison with the backup shows changes in column F alone. The comparison covers values, formats, sheet order, links, merged cells, conditional formats and filters.
8. **Recommendations.** Rebuilds `universe.json` from the layout and calc JSONs plus the repriced Main values, runs `rec.py` → `rec_<yyyymmdd>.json`, then `rec_sheet.py` + `fixlinks_all.py`. It installs only if nothing outside the Recommendations sheet changed. `rec_sheet.py` keeps your tab order and your deleted B3 line.
9. **Paper trader.** If a held name split, records it first with `paper_trader.py split`. Then runs `trader/paper_trader.py run --prices closes_<yyyymmdd>.json`. This step runs only when the Screener's F1 shows this close and step 8 installed the new sheet.

### Logs, state and backups

- `runs/<yyyymmdd>/`: `refresh.log`, per-ticker build logs, staged Screener files, the pre-run `universe.json` and `templates.json`, and `state.json`, which lets `--from` pick up where a run stopped.
- `screener/backups/`: a copy of `Screener.xlsx` before each write.
- `.refresh.lock`: prevents two refreshes from running at once. A stale lock from a crashed run is removed automatically.
- Exit codes: 0 done, 2 stopped by a safety check, 1 a script failed.

### Testing on a copy

All paths come from the script's location (and `SM_VAULT`, or `../Vault` next to `Stock Models`). To run a copy, copy `Stock Models/` (symlink `venv`) next to a copy of `Vault/Finance` and the report image, then run that copy's `refresh_daily.py`.

## Scripts

| Script | Purpose |
| --- | --- |
| `closes_live.py`, `closes.py` | Closes from the quote feed right after the bell, or the historical feed (official close, posted late at night) |
| `basic/bbuild.py`, `basic/bcalc.py` | Build a basic NPV workbook from its template; recalculate it with the `formulas` engine and cache values |
| `reprice.py` | Reprice an installed workbook in place (hand-built models) |
| `basic/bnote.py`, `bhub.py`, `brep.py` | Obsidian notes, the hub note and the Large-Cap Upside Review |
| `basic/bscreener.py`, `fixlinks_all.py` | Screener J/K/L and link caches; fix the external-link relationship Ids that openpyxl renumbers |
| `rec.py`, `rec_sheet.py` | Long/short weights from `universe.json`; write the Recommendations sheet |
| `trader/paper_trader.py` | Paper trader (see `trader/README.md`) |
