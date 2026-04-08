from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v9.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    footer_names = [n for n in zf.namelist() if n.startswith("word/footer")]
    header_names = [n for n in zf.namelist() if n.startswith("word/header")]
    footer_xml = {n: ET.fromstring(zf.read(n)) for n in footer_names}
    header_xml = {n: ET.fromstring(zf.read(n)) for n in header_names}

rel_map = {rel.attrib.get("Id"): rel.attrib.get("Target") for rel in rels}
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
            refs.append((local, ref.attrib.get(W + "type"), ref.attrib.get("{%s}id" % NS["r"]), rel_map.get(ref.attrib.get("{%s}id" % NS["r"]))))
    pg = sect.find("w:pgNumType", NS)
    sec_type = sect.find("w:type", NS)
    print(idx, repr(get_text(child)), "type=", None if sec_type is None else sec_type.attrib.get(W + "val"), "pg=", None if pg is None else pg.attrib)
    for item in refs:
        print(" ", item)

print("\n== Header Parts ==")
for name, root in header_xml.items():
    print(name, [get_text(p) for p in root.findall("w:p", NS)])

print("\n== Footer Parts ==")
for name, root in footer_xml.items():
    print(name, [get_text(p) for p in root.findall("w:p", NS)])

print("\n== Reference Block ==")
for idx in range(388, 410):
    node = children[idx - 1]
    if node.tag == W + "p":
        print(idx, get_text(node))
        if idx in (388, 389, 408, 409):
            print(ET.tostring(node, encoding="unicode"))
