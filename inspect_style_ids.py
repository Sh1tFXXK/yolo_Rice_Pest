from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]

with ZipFile(DOCX) as zf:
    styles = ET.fromstring(zf.read("word/styles.xml"))

for sid in ["9", "12", "13"]:
    style = None
    for st in styles.findall("w:style", NS):
        if st.attrib.get(W + "styleId") == sid:
            style = st
            break
    print("==", sid, "==")
    if style is None:
        print("missing")
        continue
    name = style.find("w:name", NS)
    print("name", None if name is None else name.attrib.get(W + "val"))
    ppr = style.find("w:pPr", NS)
    rpr = style.find("w:rPr", NS)
    print("ppr", ET.tostring(ppr, encoding="unicode") if ppr is not None else None)
    print("rpr", ET.tostring(rpr, encoding="unicode") if rpr is not None else None)
