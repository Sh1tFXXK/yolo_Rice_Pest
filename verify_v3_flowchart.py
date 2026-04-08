from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

DOCX_PATH = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v3.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

with ZipFile(DOCX_PATH) as zf:
    xml = zf.read("word/document.xml")
    rels_xml = zf.read("word/_rels/document.xml.rels")

root = ET.fromstring(xml)
body = root.find("w:body", NS)
children = list(body)
rels_root = ET.fromstring(rels_xml)
rels = {rel.attrib.get("Id"): rel.attrib.get("Target") for rel in rels_root}

for idx in [171, 172, 173]:
    node = children[idx - 1]
    text = "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()
    blips = node.findall(".//a:blip", NS)
    targets = []
    for b in blips:
        rid = b.attrib.get("{%s}embed" % NS["r"])
        if rid:
            targets.append(rels.get(rid, rid))
    print(idx, "text=", text, "targets=", targets)
