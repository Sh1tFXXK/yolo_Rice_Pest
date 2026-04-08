from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v2.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]

with ZipFile(DOCX_PATH) as zf:
    xml = zf.read("word/document.xml")

root = ET.fromstring(xml)
body = root.find("w:body", NS)
children = list(body)

for idx in [171, 172, 173]:
    node = children[idx - 1]
    print(f"===== {idx} =====")
    print(ET.tostring(node, encoding="unicode"))
