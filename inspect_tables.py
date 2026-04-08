import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))

body = doc.find("w:body", NS)
for idx, node in enumerate(list(body), 1):
    if node.tag != W + "tbl":
        continue
    print("== Table at body index", idx, "==")
    for row in node.findall("w:tr", NS):
        cells = []
        for cell in row.findall("w:tc", NS):
            cells.append(get_text(cell))
        print(" | ".join(cells))
    print()
