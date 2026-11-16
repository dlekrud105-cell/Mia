"""Post-process: one font (Malgun Gothic) for Latin + Korean, Korean language tag."""
import sys, zipfile, shutil, re
from lxml import etree
FONT = "Malgun Gothic"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ORDER = ["ln","noFill","solidFill","gradFill","blipFill","pattFill","grpFill","effectLst","effectDag","highlight",
         "uLnTx","uLn","uFillTx","uFill","latin","ea","cs","sym","hlinkClick","hlinkMouseOver","rtl","extLst"]
def fix_rpr(el):
    el.set("lang", "ko-KR"); el.set("altLang", "en-US")
    for tag in ("latin","ea","cs"):
        for old in el.findall(f"{{{A}}}{tag}"): el.remove(old)
        new = etree.Element(f"{{{A}}}{tag}"); new.set("typeface", FONT)
        idx = ORDER.index(tag); pos = 0
        for i, ch in enumerate(el):
            name = etree.QName(ch).localname
            if name in ORDER and ORDER.index(name) < idx: pos = i + 1
        el.insert(pos, new)
src = sys.argv[1]; tmp = src + ".tmp"
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    n = item.filename
    if n.endswith(".xml") and (n.startswith("ppt/slides/slide") or n.startswith("ppt/slideLayouts/") or n.startswith("ppt/slideMasters/") or n.startswith("ppt/charts/chart")):
        root = etree.fromstring(data)
        for el in root.iter(f"{{{A}}}rPr", f"{{{A}}}defRPr", f"{{{A}}}endParaRPr"): fix_rpr(el)
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    elif n.startswith("ppt/theme/theme"):
        root = etree.fromstring(data)
        for fnt in root.iter(f"{{{A}}}majorFont", f"{{{A}}}minorFont"):
            for tag in ("latin","ea"):
                e = fnt.find(f"{{{A}}}{tag}"); e.set("typeface", FONT)
            for f in fnt.findall(f"{{{A}}}font"):
                if f.get("script") == "Hang": f.set("typeface", "맑은 고딕")
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    zout.writestr(item, data)
zout.close(); zin.close(); shutil.move(tmp, src)
