import time, uno
from com.sun.star.connection import NoConnectException
from com.sun.star.awt.FontWeight import BOLD

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
time.sleep(1)  # let the auto-expand listener attach
sheet = doc.Sheets.getByIndex(0)

# Caption with the example formula (plain text).
cap = sheet.getCellByPosition(0, 0)
cap.setString('Example:   =STOCKDATA("AAPL", "2024-01-01", "2024-06-01")')
cap.CharWeight = BOLD
cap.CharHeight = 11

# Anchor the formula at A3 so it auto-expands the table beneath the caption.
sheet.getCellByPosition(0, 2).setFormula(
    '=STOCKDATA("AAPL";"2024-01-01";"2024-06-01")')
time.sleep(2)  # allow the listener to fetch + fill

# Bold the header row (row index 2 = sheet row 3).
header = sheet.getCellRangeByName("A3:G3")
header.CharWeight = BOLD

# Tidy column widths.
widths = [3200, 2400, 2400, 2400, 2400, 2600, 2600]
for i, w in enumerate(widths):
    sheet.Columns.getByIndex(i).Width = w

# Anchor the view at the top-left.
ctrl = doc.CurrentController
ctrl.setFirstVisibleColumn(0)
ctrl.setFirstVisibleRow(0)
ctrl.select(sheet.getCellByPosition(0, 0))
print("populated")
