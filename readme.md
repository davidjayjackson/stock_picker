# StockData — a LibreOffice Calc extension

A LibreOffice Calc extension that adds a `STOCKDATA` worksheet function, the
Calc equivalent of Excel's `STOCKHISTORY`. It pulls historical daily stock
data from Yahoo Finance.

```
=STOCKDATA("AAPL", "2024-01-01", "2024-03-01")
```

The function takes three arguments — **stock symbol**, **start date**, and
**end date** — and returns all the daily data columns as a table. Dates can be
text (`"2024-01-01"`), date cells, or live formulas like `TODAY()`:

```
=STOCKDATA("AAPL", "2024-01-01", TODAY())
```


| Date | Open | High | Low | Close | Adj Close | Volume |
|------|------|------|-----|-------|-----------|--------|

## Getting the whole table

LibreOffice does not auto-spill formula arrays the way Excel 365 does, so the
extension offers two ways to land the full result:

- **Auto-expand (default).** Type `=STOCKDATA(...)` in a single empty cell and
  press **Enter** — the surrounding cells fill in automatically with the table
  as static values. (One-shot fill; re-enter the formula to refresh. Takes
  effect in documents opened after the extension is installed.)
- **Live array formula.** Select a range 7 columns wide and tall enough for the
  data, type the formula, and press **Ctrl+Shift+Enter** for a reactive result.

## Project layout

```
stock_picker/
├── src/
│   ├── idl/XStockData.idl      UNO interface, compiled to StockData.rdb
│   └── python/stockdata.py     add-in function + auto-expand listener/job
├── oxt/                        extension metadata & registration
│   ├── description.xml
│   ├── CalcAddIns.xcu          Function Wizard registration
│   ├── Jobs.xcu                document-open Job (auto-expand)
│   └── META-INF/manifest.xml
├── build.ps1 / build.sh        compile the IDL and pack dist/StockData.oxt
├── tools/                      headless UNO tests
└── BUILD.md                    build, install, and usage guide
```

## Building & installing

See **[BUILD.md](BUILD.md)**. In short: install the LibreOffice SDK, run
`build.ps1` (Windows) or `build.sh` (Linux/macOS) to produce
`dist/StockData.oxt`, then install it with `unopkg add`. The add-in itself uses
only the Python standard library — no third-party packages.

## How it works

- `XStockData.idl` declares the UNO function signature; the build compiles it
  into the `StockData.rdb` type library shipped inside the `.oxt`.
- `stockdata.py` implements that interface plus `com.sun.star.sheet.AddIn`, and
  fetches data from Yahoo's public chart endpoint
  (`https://query1.finance.yahoo.com/v8/finance/chart/<symbol>`).
- A document-open Job attaches a modify listener that powers auto-expand.
