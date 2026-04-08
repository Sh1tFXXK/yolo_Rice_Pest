from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]

def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()

with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
body = doc.find("w:body", NS)
children = list(body)
for idx in range(max(1, len(children)-25), len(children)+1):
    node = children[idx-1]
    if node.tag == W + "p":
        print(idx, repr(get_text(node)))
    else:
        print(idx, node.tag)
