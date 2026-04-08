from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


patterns = [
    re.compile(r"^\d+[.、]"),
    re.compile(r"^[（(]\d+[)）]"),
    re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]"),
]

with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))

body = doc.find("w:body", NS)
for idx, node in enumerate(list(body), 1):
    if node.tag != W + "p":
        continue
    text = get_text(node)
    if text and any(p.search(text) for p in patterns):
        print(idx, text)
