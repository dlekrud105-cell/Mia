"""Export PPTX to PDF via LibreOffice UNO, disabling Asian/Latin auto-spacing
(LibreOffice inserts visible gaps around punctuation in Korean text otherwise)."""
import sys, os, time, subprocess, tempfile, uno
from com.sun.star.beans import PropertyValue
sys.path.insert(0, "/root/.claude/skills/synced/eb7b2f3f-445a-4846-acbc-22136f2f01b5_65af6b0e-d336-4b83-bddb-739ee7f577d1/pptx/scripts")
from office.soffice import get_soffice_env
src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
prof = tempfile.mkdtemp()
p = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{prof}",
                      "--accept=socket,host=127.0.0.1,port=2002;urp;"], env=get_soffice_env(),
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = uno.getComponentContext()
res = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
for _ in range(60):
    try:
        rctx = res.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
desk = rctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", rctx)
def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
def fix(shape):
    if shape.supportsService("com.sun.star.drawing.GroupShape"):
        for i in range(shape.getCount()): fix(shape.getByIndex(i))
        return
    if shape.ShapeType == "com.sun.star.drawing.TableShape":
        m = shape.Model
        for r in range(m.Rows.Count):
            for c in range(m.Columns.Count):
                cell = m.getCellByPosition(c, r)
                try:
                    cur = cell.createTextCursor(); cur.gotoStart(False); cur.gotoEnd(True)
                    cur.ParaIsCharacterDistance = False
                except Exception as e: print("cell cursor:", e)
                en = cell.createEnumeration()
                while en.hasMoreElements():
                    para = en.nextElement()
                    try: para.ParaIsCharacterDistance = False
                    except Exception: pass
        return
    if hasattr(shape, "getText"): fix_text(shape)
def fix_text(obj):
    try:
        txt = obj.getText() if hasattr(obj, "getText") else obj
        en = txt.createEnumeration()
        while en.hasMoreElements():
            para = en.nextElement()
            try: para.ParaIsCharacterDistance = False
            except Exception: pass
    except Exception: pass
pages = doc.getDrawPages()
for i in range(pages.Count):
    pg = pages.getByIndex(i)
    for j in range(pg.Count): fix(pg.getByIndex(j))
doc.storeToURL(uno.systemPathToFileUrl(dst), (pv("FilterName", "impress_pdf_Export"),))
doc.close(True)
try: desk.terminate()
except Exception: pass
p.wait(timeout=30)
