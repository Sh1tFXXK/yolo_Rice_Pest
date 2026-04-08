from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))

rel_map = {rel.attrib.get("Id"): (rel.attrib.get("Type"), rel.attrib.get("Target")) for rel in rels}
body = doc.find("w:body", NS)
children = list(body)

print("== Sections ==")
for idx, child in enumerate(children, 1):
    sect = None
    if child.tag == W + "p":
        ppr = child.find("w:pPr", NS)
        sect = None if ppr is None else ppr.find("w:sectPr", NS)
    elif child.tag == W + "sectPr":
        sect = child
    if sect is None:
        continue
    refs = []
    for ref in sect:
        local = ref.tag.split("}")[-1]
        if local in {"headerReference", "footerReference"}:
            rid = ref.attrib.get("{%s}id" % NS["r"])
            refs.append((local, ref.attrib.get(W + "type"), rid, rel_map.get(rid)))
    pg = sect.find("w:pgNumType", NS)
    sec_type = sect.find("w:type", NS)
    print("index", idx, "text", repr(get_text(child)[:50]), "type", None if sec_type is None else sec_type.attrib.get(W + "val"), "pgFmt", None if pg is None else pg.attrib.get(W + "fmt"))
    for item in refs:
        print(" ", item)

print("\n== Footer files ==")
for rid, info in rel_map.items():
    if info[0] and info[0].endswith("/footer"):
        print(rid, info)
    if info[0] and info[0].endswith("/header"):
        print(rid, info)

print("\n== References tail ==")
for idx in range(388, len(children) + 1):
    node = children[idx - 1]
    if node.tag == W + "p":
        text = get_text(node)
        if text:
            print(idx, text)
