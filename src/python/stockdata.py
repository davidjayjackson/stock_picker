# -*- coding: utf-8 -*-
"""
LibreOffice Calc add-in: STOCKDATA().

Implements the IDL interface com.davidjackson.stockpicker.XStockData and the
standard com.sun.star.sheet.AddIn service so that Calc exposes the cell
function:

    =STOCKDATA(symbol; start_date; end_date)

The function returns a 2D array (header row + one row per trading day) of
historical daily data pulled from Yahoo Finance's public chart endpoint.

Two ways to get the whole table into the sheet:

* Auto-expand (default): type =STOCKDATA(...) in a single cell and press
  Enter. A document modify listener (see AutoExpandListener below) fills the
  surrounding cells with the table as static values. One-shot, no recalc.
* Live array formula: select a range, type the formula and press
  Ctrl+Shift+Enter for a reactive result.
"""

import datetime
import json
import re
import unohelper
import uno

from urllib.parse import quote
from urllib.request import Request, urlopen

from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XServiceName, XServiceInfo
from com.sun.star.task import XJob
from com.sun.star.util import XModifyListener

# The custom interface comes from this extension's own type library
# (StockData.rdb). It resolves at load time because the office process has the
# installed extension's types available. Inheriting it is what lets Calc
# introspect the stockdata() signature (and thus its 3 parameters).
from com.davidjackson.stockpicker import XStockData

# --- identity -----------------------------------------------------------------

# The implementation name doubles as the add-in's service name (getServiceName)
# and the CalcAddIns.xcu AddInInfo node name; Calc keys the add-in's functions
# off this string, so the three must stay identical.
IMPL_NAME = "com.davidjackson.stockpicker.StockDataImpl"
ADDIN_SERVICE = "com.sun.star.sheet.AddIn"

# Programmatic name of the single function we expose. Must match the node
# names used in CalcAddIns.xcu and the method name on the IDL interface.
FUNC = "stockdata"
CATEGORY = "Add-In"

COLUMNS = ("Date", "Open", "High", "Low", "Close", "Adj Close", "Volume")

# Display metadata, also mirrored declaratively in CalcAddIns.xcu.
_ARG_NAMES = ("Symbol", "Start Date", "End Date")
_ARG_DESCS = (
    "Ticker symbol, e.g. \"AAPL\".",
    "First day to include (YYYY-MM-DD or a date cell).",
    "Last day to include, inclusive (YYYY-MM-DD or a date cell).",
)

# LibreOffice/Excel date serials count days from 1899-12-30.
_SPREADSHEET_EPOCH = datetime.date(1899, 12, 30)


# --- date handling ------------------------------------------------------------

def _parse_date(value):
    """Accept an ISO/US date string or a spreadsheet serial number."""
    if value is None:
        raise ValueError("missing date")

    text = str(value).strip()
    if text == "":
        raise ValueError("missing date")

    # A bare number is a spreadsheet date serial.
    try:
        serial = float(text)
        return _SPREADSHEET_EPOCH + datetime.timedelta(days=int(serial))
    except ValueError:
        pass

    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    raise ValueError("unrecognized date: %r" % value)


def _to_epoch(date_obj):
    return int(datetime.datetime(
        date_obj.year, date_obj.month, date_obj.day,
        tzinfo=datetime.timezone.utc).timestamp())


# --- data fetching ------------------------------------------------------------

def _num(value):
    """Normalize a JSON number; Yahoo uses null for gaps/holidays."""
    return "" if value is None else value


def _fetch(symbol, start_date, end_date):
    start = _parse_date(start_date)
    end = _parse_date(end_date)
    if end < start:
        raise ValueError("end date is before start date")

    period1 = _to_epoch(start)
    # period2 is exclusive; add a day so the end date is included.
    period2 = _to_epoch(end) + 86400

    url = (
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        + quote(str(symbol).strip())
        + "?period1=%d&period2=%d&interval=1d&events=div%%2Csplit"
        % (period1, period2)
    )

    req = Request(url, headers={
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        ),
        "Accept": "application/json",
    })
    with urlopen(req, timeout=30) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    chart = payload.get("chart") or {}
    err = chart.get("error")
    if err:
        raise RuntimeError(err.get("description") or err.get("code") or str(err))

    results = chart.get("result") or []
    if not results:
        raise RuntimeError("no data returned for %r" % symbol)

    result = results[0]
    timestamps = result.get("timestamp") or []
    indicators = result.get("indicators") or {}
    quote_block = (indicators.get("quote") or [{}])[0]
    adj_block = (indicators.get("adjclose") or [{}])[0]
    adjclose = adj_block.get("adjclose")

    opens = quote_block.get("open") or []
    highs = quote_block.get("high") or []
    lows = quote_block.get("low") or []
    closes = quote_block.get("close") or []
    volumes = quote_block.get("volume") or []

    rows = [COLUMNS]
    for i, ts in enumerate(timestamps):
        day = datetime.datetime.fromtimestamp(
            ts, datetime.timezone.utc).strftime("%Y-%m-%d")
        adj = adjclose[i] if adjclose and i < len(adjclose) else (
            closes[i] if i < len(closes) else None)
        rows.append((
            day,
            _num(opens[i]) if i < len(opens) else "",
            _num(highs[i]) if i < len(highs) else "",
            _num(lows[i]) if i < len(lows) else "",
            _num(closes[i]) if i < len(closes) else "",
            _num(adj),
            _num(volumes[i]) if i < len(volumes) else "",
        ))

    if len(rows) == 1:
        raise RuntimeError("no trading days in the requested range")

    return tuple(rows)


# --- UNO component ------------------------------------------------------------

class StockDataImpl(unohelper.Base, XStockData, XAddIn,
                    XServiceName, XServiceInfo):
    """The add-in implementation.

    Inheriting ``XStockData`` (resolved from this extension's own type
    library at load time) is what lets Calc introspect the ``stockdata``
    method's signature, so the function is registered with its three
    arguments. ``XAddIn`` supplies the Function Wizard display names and
    ``XServiceName``/``XServiceInfo`` identify the component.
    """

    def __init__(self, ctx):
        self.ctx = ctx
        self._locale = None

    # -- XStockData --------------------------------------------------------
    def stockdata(self, symbol, startDate, endDate):
        try:
            return _fetch(symbol, startDate, endDate)
        except Exception as exc:  # surface the error inside the cell
            return (("STOCKDATA error: %s" % exc,),)

    # -- XServiceName ------------------------------------------------------
    def getServiceName(self):
        # Calc keys the add-in's functions by this name, and matches the
        # CalcAddIns.xcu AddInInfo node against the *implementation* name.
        # Returning the implementation name here makes the two keys line up
        # so the .xcu descriptions/parameter help merge onto the function.
        return IMPL_NAME

    # -- XServiceInfo ------------------------------------------------------
    def getImplementationName(self):
        return IMPL_NAME

    def supportsService(self, name):
        return name == ADDIN_SERVICE

    def getSupportedServiceNames(self):
        return (ADDIN_SERVICE,)

    # -- XAddIn ------------------------------------------------------------
    def getProgrammaticFuntionName(self, displayName):  # historical typo
        return FUNC if displayName.upper() == "STOCKDATA" else ""

    def getDisplayFunctionName(self, programmaticName):
        return "STOCKDATA" if programmaticName == FUNC else ""

    def getDisplayArgumentName(self, programmaticName, argument):
        if programmaticName == FUNC and 0 <= argument < len(_ARG_NAMES):
            return _ARG_NAMES[argument]
        return ""

    def getDisplayArgumentDescription(self, programmaticName, argument):
        if programmaticName == FUNC and 0 <= argument < len(_ARG_DESCS):
            return _ARG_DESCS[argument]
        return ""

    def getProgrammaticCategoryName(self, programmaticName):
        return CATEGORY

    def getDisplayCategoryName(self, programmaticName):
        return CATEGORY

    # -- XLocalizable (part of the AddIn service) --------------------------
    def setLocale(self, locale):
        self._locale = locale

    def getLocale(self):
        if self._locale is not None:
            return self._locale
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = "en"
        locale.Country = "US"
        return locale


# --- auto-expand --------------------------------------------------------------
#
# A worksheet function cannot write to neighbouring cells, so a plain
# =STOCKDATA(...) in a single cell only shows the table's top-left value
# ("Date"). To make it "spill" automatically we attach a modify listener to
# each spreadsheet (via a document-open Job). When the listener sees a fresh
# single-cell STOCKDATA formula it fetches the data and writes the whole table
# as static values in place. This is a one-shot fill; the live array-formula
# usage (Ctrl+Shift+Enter over a range) still works for a reactive result.

JOB_IMPL_NAME = "com.davidjackson.stockpicker.AutoExpandJob"
JOB_SERVICE = "com.davidjackson.stockpicker.AutoExpandJob"
SPREADSHEET_SERVICE = "com.sun.star.sheet.SpreadsheetDocument"

# Documents we have already wired up, keyed by RuntimeUID.
_ATTACHED = set()

_CALL_RE = re.compile(r"STOCKDATA\s*\((.*)\)\s*$", re.IGNORECASE | re.DOTALL)


def _parse_call(formula):
    """Pull (symbol, start, end) out of a STOCKDATA formula string."""
    match = _CALL_RE.search(formula)
    if not match:
        return None
    parts = re.split(r"[;,]", match.group(1))
    vals = [p.strip().strip('"').strip() for p in parts]
    if len(vals) < 3 or not vals[0]:
        return None
    return vals[0], vals[1], vals[2]


def _coerce(value):
    """Make a fetched value safe for XCellRange.setDataArray (string/double)."""
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    return "" if value is None else str(value)


class AutoExpandListener(unohelper.Base, XModifyListener):
    """Expands single-cell STOCKDATA formulas on the active sheet in place."""

    def __init__(self, doc):
        self.doc = doc
        self._busy = False

    def modified(self, event):
        if self._busy:
            return
        try:
            self._expand()
        except Exception:
            # Never let a listener error disrupt the user's editing.
            pass

    def disposing(self, event):
        try:
            _ATTACHED.discard(self.doc.RuntimeUID)
        except Exception:
            pass

    def _expand(self):
        try:
            sheet = self.doc.CurrentController.ActiveSheet
        except Exception:
            return

        formula_flag = uno.getConstantByName(
            "com.sun.star.sheet.CellFlags.FORMULA")
        ranges = sheet.queryContentCells(formula_flag)
        cells = ranges.getCells()
        if not cells.hasElements():
            return

        targets = []
        enum = cells.createEnumeration()
        while enum.hasMoreElements():
            cell = enum.nextElement()
            if "STOCKDATA(" not in cell.getFormula().upper():
                continue
            addr = cell.CellAddress
            # Only act on a fresh single-cell entry: skip if a neighbour is
            # already occupied (an array formula or a previous expansion).
            right = sheet.getCellByPosition(addr.Column + 1, addr.Row)
            below = sheet.getCellByPosition(addr.Column, addr.Row + 1)
            if right.getString() or below.getString():
                continue
            parsed = _parse_call(cell.getFormula())
            if parsed:
                targets.append((addr, parsed))

        if not targets:
            return

        self._busy = True
        try:
            for addr, (symbol, start, end) in targets:
                try:
                    rows = _fetch(symbol, start, end)
                except Exception as exc:
                    rows = (("STOCKDATA error: %s" % exc,),)
                ncols = max(len(r) for r in rows)
                data = tuple(
                    tuple([_coerce(v) for v in r] + [""] * (ncols - len(r)))
                    for r in rows
                )
                target = sheet.getCellRangeByPosition(
                    addr.Column, addr.Row,
                    addr.Column + ncols - 1, addr.Row + len(data) - 1)
                target.setDataArray(data)
        finally:
            self._busy = False


class AutoExpandJob(unohelper.Base, XJob, XServiceInfo):
    """Runs on document open/new and attaches the modify listener once."""

    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, arguments):
        model = self._find_model(arguments)
        if model is None:
            return None
        try:
            if not model.supportsService(SPREADSHEET_SERVICE):
                return None
            uid = model.RuntimeUID
        except Exception:
            return None
        if uid in _ATTACHED:
            return None
        try:
            model.addModifyListener(AutoExpandListener(model))
            _ATTACHED.add(uid)
        except Exception:
            pass
        return None

    @staticmethod
    def _find_model(arguments):
        for nv in arguments or ():
            if nv.Name == "Environment":
                for env in nv.Value or ():
                    if env.Name == "Model":
                        return env.Value
        return None

    # -- XServiceInfo ------------------------------------------------------
    def getImplementationName(self):
        return JOB_IMPL_NAME

    def supportsService(self, name):
        return name == JOB_SERVICE

    def getSupportedServiceNames(self):
        return (JOB_SERVICE,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    StockDataImpl, IMPL_NAME, (ADDIN_SERVICE,),
)
g_ImplementationHelper.addImplementation(
    AutoExpandJob, JOB_IMPL_NAME, (JOB_SERVICE,),
)
