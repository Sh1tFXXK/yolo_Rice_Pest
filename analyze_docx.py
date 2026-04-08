from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path
import re

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_working.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
KEYWORDS = ["未", "拟", "计划", "将", "准备", "预期", "后续", "待", "预计", "本研究拟", "本系统拟", "将采用"]


def text_from_node(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


OUT_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_analysis.txt")

with ZipFile(DOCX_PATH) as zf:
    document_xml = zf.read("word/document.xml")
    rels_xml = zf.read("word/_rels/document.xml.rels")
    names = zf.namelist()

root = ET.fromstring(document_xml)
body = root.find("w:body", NS)
rels_root = ET.fromstring(rels_xml)
rels = {}
for rel in rels_root:
    rid = rel.attrib.get("Id")
    target = rel.attrib.get("Target")
    if rid and target:
        rels[rid] = target

lines = []
lines.append("== Keywords ==")
for idx, child in enumerate(body, 1):
    text = text_from_node(child)
    if text and any(k in text for k in KEYWORDS):
        lines.append(f"{idx}: {text}")

lines.append("")
lines.append("== Captions ==")
for idx, child in enumerate(body, 1):
    text = text_from_node(child)
    if text and (re.search(r"^图\s*\d", text) or re.search(r"^表\s*\d", text) or "流程图" in text):
        lines.append(f"{idx}: {text}")

lines.append("")
lines.append("== Drawings ==")
for idx, child in enumerate(body, 1):
    if child.findall(".//w:drawing", NS):
        text = text_from_node(child)
        blips = child.findall(".//a:blip", NS)
        targets = []
        for blip in blips:
            rid = blip.attrib.get(f"{{{NS['r']}}}embed")
            if rid:
                targets.append(rels.get(rid, rid))
        lines.append(f"{idx}: text={text!r} media={targets}")

lines.append("")
lines.append("== Media Files ==")
for name in names:
    if name.startswith("word/media/"):
        lines.append(name)

OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
print(str(OUT_PATH))
