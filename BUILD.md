# Building & installing the StockData add-in

`STOCKDATA()` is a LibreOffice Calc add-in function, the Calc equivalent of
Excel's `STOCKHISTORY`:

```
=STOCKDATA("AAPL"; "2024-01-01"; "2024-03-01")
```

It returns a 2D table — a header row plus one row per trading day with columns
**Date, Open, High, Low, Close, Adj Close, Volume** — fetched live from Yahoo
Finance.

## Directory layout

```
stock_picker/
├── src/
│   ├── idl/XStockData.idl      # UNO interface (compiled to StockData.rdb)
│   └── python/stockdata.py     # the add-in implementation
├── oxt/
│   ├── description.xml         # extension metadata
│   ├── CalcAddIns.xcu          # function/argument registration & help text
│   ├── description/desc_en.txt # extension blurb
│   └── META-INF/manifest.xml   # package manifest
├── build.ps1                   # Windows build
├── build.sh                    # Linux/macOS build
└── dist/StockData.oxt          # produced by the build
```

## Prerequisites

- **LibreOffice** (provides the runtime, `unopkg`, and `types.rdb`).
- **LibreOffice SDK** — only needed to compile the IDL into a type library.
  Install the SDK package that matches your LibreOffice version. The build
  prefers the modern `unoidl-write` tool (shipped in `program/`); it falls
  back to the SDK's `idlc` + `regmerge` if present.

The add-in itself needs **no third-party Python packages** — it uses only the
standard library (`urllib`, `json`, `datetime`) against the Python interpreter
bundled with LibreOffice.

## Build

Windows (PowerShell):

```powershell
.\build.ps1
# or point it explicitly at your install:
.\build.ps1 -OfficeDir "C:\Program Files\LibreOffice\program" `
            -SdkDir    "C:\Program Files\LibreOffice\sdk"
```

Linux/macOS:

```bash
LO_PROGRAM=/usr/lib/libreoffice/program ./build.sh
```

Both produce `dist/StockData.oxt`.

## Install

```powershell
& "C:\Program Files\LibreOffice\program\unopkg.exe" add --force dist\StockData.oxt
```

Or in the GUI: **Tools ▸ Extension Manager ▸ Add…**, pick the `.oxt`, then
restart LibreOffice. To remove:

```powershell
unopkg remove com.davidjackson.stockpicker
```

## Use

In any Calc cell:

```
=STOCKDATA("MSFT"; "2024-01-01"; "2024-02-01")
```

Dates may be ISO strings (`YYYY-MM-DD`), common regional formats, references
to date cells, or live formulas such as `TODAY()` or `DATE(2026,1,1)` — these
work in both the live array-formula and auto-expand forms. Errors (bad symbol,
no network, empty range) appear as a single `STOCKDATA error: …` cell.

```
=STOCKDATA("AAPL", "2024-01-01", TODAY())
```

### Getting the whole table

`STOCKDATA` returns a table (header + one row per trading day). LibreOffice
does not auto-spill formula arrays the way Excel 365 does, so there are two
ways to land the full result:

1. **Auto-expand (default).** Type `=STOCKDATA(...)` in a single empty cell and
   press **Enter**. The extension's modify listener fetches the data and fills
   the cells around the anchor with the table. This is a **one-shot fill**: the
   anchor formula is replaced by static values, so the data does not refresh on
   recalc — re-enter the formula to refresh. The anchor cell should be in an
   empty area (the listener skips if a neighbouring cell is already occupied).
2. **Live array formula.** Select a range 7 columns wide and tall enough for the
   data, type the formula, and press **Ctrl+Shift+Enter**. This stays reactive
   and recalculates, but you must size the range yourself.

Auto-expand is wired up by a document-open Job (`Jobs.xcu` →
`AutoExpandJob`) that attaches `AutoExpandListener` to each spreadsheet. It
takes effect on documents opened/created **after** the extension is installed,
so reopen any document that was already open at install time.

## Testing

`tools/test_addin.py` drives a headless LibreOffice over a UNO socket, calls
`STOCKDATA` from a real array formula, and prints the spilled result. Run it
with LibreOffice's bundled Python:

```powershell
# 1. start a headless office with a socket
Start-Process "C:\Program Files\LibreOffice\program\soffice.exe" `
  -ArgumentList '--headless','--norestore','--invisible','--nodefault','--nologo',`
                '--accept=socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'
# 2. run the test
& "C:\Program Files\LibreOffice\program\python.exe" tools\test_addin.py
```

Expected output is a header row plus one row per trading day in the range.

## Known limitation: Function Wizard descriptions

The function name (`STOCKDATA`) and the argument names (`Symbol`, `Start
Date`, `End Date`) show correctly in the Function Wizard. The longer
*description* help texts may render as `###`.

This was chased down thoroughly: the `CalcAddIns.xcu` is structured exactly
like the documented working examples, and the descriptions are confirmed
present in LibreOffice's live configuration registry after install. In this
LibreOffice version, however, the function manager does not merge those
declarative descriptions onto a Python add-in function discovered via
interface introspection. It is purely cosmetic — the function and all its
arguments work normally.

## Publishing to extensions.libreoffice.org

The `.oxt` produced by the build **is** the distributable package — that is the
file you upload. It already carries everything the gallery needs:

- `description.xml` — identifier, version (`1.0.0`), display name, publisher,
  icon, and a click-through MIT license (`registration/simple-license`) shown
  on install.
- `icons/icon.png` — 42×42 icon for the Extension Manager and the listing.
- `registration/LICENSE`, `description/desc_en.txt` — license text and blurb.

Steps:

1. Run `.\build.ps1` to produce `dist/StockData.oxt`.
2. Sign in at <https://extensions.libreoffice.org/> and add a new project:
   title, summary, full description, and a category (e.g. *Calc* +
   *Business/Finance*).
3. Add a release: upload the `.oxt`, set the **compatible LibreOffice
   versions**, choose the **MIT** license, and attach **screenshots** (the
   spilled-table view works well).
4. Submit — listings go through moderator review before they appear publicly.

Bump `<version>` in `oxt/description.xml` for each release. To offer in-app
update notifications, host an update feed and add an `<update-information>`
element pointing at it.

## How it works

- `XStockData.idl` declares the UNO interface `stockdata(string, string,
  string) -> sequence<sequence<any>>`. The build compiles it into
  `StockData.rdb`, the type library shipped inside the `.oxt`.
- `stockdata.py` implements that interface plus the standard
  `com.sun.star.sheet.AddIn` service, and registers itself with
  `g_ImplementationHelper`. Python-UNO dispatches the `stockdata` call by
  method name.
- `CalcAddIns.xcu` registers the function, its category, and the argument
  names/help shown in the Function Wizard. The parameter node names
  (`symbol`, `startDate`, `endDate`) must match the IDL argument names.
- Data comes from Yahoo's public chart endpoint
  `https://query1.finance.yahoo.com/v8/finance/chart/<symbol>`.
