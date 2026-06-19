# Run with LibreOffice's bundled python against a headless soffice on port 2002:
#   "C:\Program Files\LibreOffice\program\python.exe" tools\test_addin.py
import sys
import time
import uno
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    last = None
    for _ in range(30):
        try:
            return resolver.resolve(url)
        except NoConnectException as exc:
            last = exc
            time.sleep(1)
    raise last


def main():
    ctx = connect()
    smgr = ctx.ServiceManager

    # 1) Find how the add-in function is registered (search descriptions).
    fd = smgr.createInstanceWithContext(
        "com.sun.star.sheet.FunctionDescriptions", ctx)
    print("--- matching functions ---")
    for i in range(fd.Count):
        props = fd.getByIndex(i)
        d = {p.Name: p.Value for p in props}
        name = d.get("Name", "")
        if "STOCK" in str(name).upper():
            print("Name=%r  Id=%s" % (name, d.get("Id")))

    # 2) Real end-to-end test: put the formula in a cell and read it back.
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
    doc = desktop.loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, (hidden,))
    try:
        sheet = doc.Sheets.getByIndex(0)
        rng = sheet.getCellRangeByName("A1:G6")
        rng.setArrayFormula('=STOCKDATA("AAPL";"2024-01-02";"2024-01-05")')
        doc.calculateAll()

        a1 = sheet.getCellByPosition(0, 0)
        print("--- A1 ---")
        print("formula:", a1.getFormula())
        print("string :", a1.getString())
        print("--- spilled block (A1:G6) ---")
        for r in range(6):
            vals = []
            for c in range(7):
                vals.append(sheet.getCellByPosition(c, r).getString())
            if any(vals):
                print(" | ".join(vals))
    finally:
        doc.close(False)


if __name__ == "__main__":
    sys.exit(main())
