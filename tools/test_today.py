import time, uno
from com.sun.star.connection import NoConnectException

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
time.sleep(1)
sheet = doc.Sheets.getByIndex(0)

print("=== A) LIVE array formula with TODAY() ===")
rng = sheet.getCellRangeByName("A1:G5")
rng.setArrayFormula('=STOCKDATA("IBM";"2026-06-10";TODAY())')
doc.calculateAll()
for r in range(5):
    vals = [sheet.getCellByPosition(c, r).getString() for c in range(7)]
    if any(vals):
        print(" | ".join(vals))

print("\n=== B) AUTO-EXPAND single cell with TODAY() ===")
sheet.getCellByPosition(0, 10).setFormula(
    '=STOCKDATA("IBM";"2026-06-10";TODAY())')
time.sleep(2)
for r in range(10, 16):
    vals = [sheet.getCellByPosition(c, r).getString() for c in range(7)]
    if any(vals):
        print(" | ".join(vals))

doc.close(False)
