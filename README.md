# Stock-Models

Valuation models for the biotech and pharma names on the screener.

- `Screener.xlsx`: **Pharma** tab (price, market cap, EV, model value, upside, model type, next event, notes), **Recommendations** (model-weighted 60/40 long/short book), **Energy**.
- `Pharma/<TICKER>.xlsx`: one workbook per company. Main sheet (price is pulled from the Screener), Yearly Model with every fiscal year of SEC data, plus either:
  - **DCF**: 10-year free-cash-flow model with bear/base/bull cases and sensitivity tables (an Options sheet for the reviewed large caps), or
  - **rNPV**: program-level risk-adjusted NPV for the PDUFA-event names, with an Options sheet, or
  - **Net cash**: for pre-revenue companies.
- `.model-tools/`: the Python pipeline that builds, prices and checks the models from SEC XBRL data, Nasdaq/Cboe prices and option chains.

Model output, not investment advice.
