from collections import Counter
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v6.docx")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    styles = ET.fromstring(zf.read("word/styles.xml"))

style_names = {}
for style in styles.findall("w:style", NS):
    style_id = style.attrib.get(W + "styleId")
    name = style.find("w:name", NS)
    style_names[style_id] = name.attrib.get(W + "val") if name is not None else style_id

body = doc.find("w:body", NS)
children = list(body)

counts = Counter()
for child in children:
    if child.tag != W + "p":
        continue
    text = get_text(child)
    if not text:
        continue
    ppr = child.find("w:pPr", NS)
    style_id = "(none)"
    if ppr is not None:
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is not None:
            style_id = pstyle.attrib.get(W + "val")
    counts[style_id] += 1

print("== Paragraph Style Counts ==")
for style_id, num in counts.most_common(30):
    print(style_id, style_names.get(style_id, style_id), num)

print("\n== Candidate Headings/Captions ==")
patterns = [
    re.compile(r"^第[一二三四五六七八九十]+章"),
    re.compile(r"^\d+(\.\d+)*\s"),
    re.compile(r"^摘要$"),
    re.compile(r"^Abstract$"),
    re.compile(r"^参考文献$"),
    re.compile(r"^图\s*\d"),
    re.compile(r"^表\s*\d"),
]
for idx, child in enumerate(children, 1):
    if child.tag != W + "p":
        continue
    text = get_text(child)
    if not text:
        continue
    if not any(p.search(text) for p in patterns):
        continue
    ppr = child.find("w:pPr", NS)
    style_id = "(none)"
    if ppr is not None:
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is not None:
            style_id = pstyle.attrib.get(W + "val")
    print(idx, style_id, style_names.get(style_id, style_id), text)

print("\n== Tables ==")
for idx, child in enumerate(children, 1):
    if child.tag == W + "tbl":
        print("table body index", idx)
