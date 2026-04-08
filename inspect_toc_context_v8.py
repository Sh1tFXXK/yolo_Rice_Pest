from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))

body = doc.find("w:body", NS)
children = list(body)

for idx in range(50, 58):
    node = children[idx - 1]
    print("====", idx, node.tag)
    print(ET.tostring(node, encoding="unicode"))
