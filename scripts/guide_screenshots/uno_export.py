"""Export chosen sheets (or ranges) of an xlsx as one PDF each.

Usage: python uno_export.py <abs-xlsx> <outdir> "Sheet[:A1:N12],Sheet2,..."
Sheet names may use ~ for spaces. An optional :RANGE exports just that
range (tight screenshots for tall input sheets).
"""
import sys

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


src, outdir, wanted = sys.argv[1], sys.argv[2], sys.argv[3]
unprotect_pw = sys.argv[4] if len(sys.argv) > 4 else None

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", localContext)
ctx = resolver.resolve(
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", ctx)

doc = desktop.loadComponentFromURL(
    "file://" + src, "_blank", 0, (prop("Hidden", True),))
doc.calculateAll()                     # xlsx stores no cached formula results

# comments would print as indicator icons — drop them all for clean shots
# (masked builds protect every sheet, which blocks the removal: unprotect
# first with the derived sheet password)
for i in range(doc.Sheets.Count):
    sh = doc.Sheets.getByIndex(i)
    if unprotect_pw and sh.isProtected():
        sh.unprotect(unprotect_pw)
    ann = sh.Annotations
    while ann.Count:
        ann.removeByIndex(0)

# clean page setup: landscape, fit-to-width, no header/footer, slim margins
styles = doc.StyleFamilies.getByName("PageStyles")
for i in range(styles.Count):
    st = styles.getByIndex(i)
    st.IsLandscape = True
    st.ScaleToPagesX = 1
    st.HeaderIsOn = False
    st.FooterIsOn = False
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(st, m, 300)            # 3 mm

for spec in wanted.split(","):
    spec = spec.replace("~", " ")
    if ":" in spec:
        name, rng = spec.split(":", 1)
    else:
        name, rng = spec, None
    sheet = doc.Sheets.getByName(name)
    sheet.IsVisible = True
    target = sheet.getCellRangeByName(rng) if rng else sheet
    filter_data = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        (prop("Selection", target),))
    safe = name.lower().replace(" ", "-").replace("_", "-")
    doc.storeToURL(
        "file://%s/%s.pdf" % (outdir, safe),
        (prop("FilterName", "calc_pdf_Export"),
         prop("FilterData", filter_data)))
    print("exported", spec)

doc.close(False)
