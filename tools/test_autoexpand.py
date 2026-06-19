# Verifies auto-expand: opening a Calc doc should attach the listener (via the
# document-open Job), and entering a single-cell STOCKDATA formula should spill
# the full table automatically (no Ctrl+Shift+Enter).
import time, uno
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue

def connect():
    local = uno.getComponentContext()
    r = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(30):
        try:
            return r.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except NoConnectException:
            time.sleep(1)
    raise SystemExit("no connection")

ctx = connect()
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
time.sleep(1)  # let the document-open Job attach the listener

sheet = doc.Sheets.getByIndex(0)
# Enter as a PLAIN single-cell formula (what the user did).
sheet.getCellByPosition(0, 0).setFormula(
    '=STOCKDATA("IBM";"2026-01-01";"2026-06-01")')
time.sleep(2)  # allow the listener to fetch + fill

print("A1 formula:", repr(sheet.getCellByPosition(0, 0).getFormula()))
print("--- first rows of the spilled block ---")
nonempty = 0
for r in range(6):
    vals = [sheet.getCellByPosition(c, r).getString() for c in range(7)]
    if any(vals):
        nonempty += 1
        print(" | ".join(vals))
# count total filled rows in column A
filled = 0
for r in range(400):
    if sheet.getCellByPosition(0, r).getString():
        filled += 1
print("total filled rows in col A:", filled)
doc.close(False)
