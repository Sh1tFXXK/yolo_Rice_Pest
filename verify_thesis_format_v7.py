from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v7.docx")
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{%s}" % NS["w"]


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def show_spacing(node):
    ppr = node.find("w:pPr", NS)
    if ppr is None:
        return {}
    spacing = ppr.find("w:spacing", NS)
    ind = ppr.find("w:ind", NS)
    jc = ppr.find("w:jc", NS)
    out = {}
    if spacing is not None:
        out["line"] = spacing.attrib.get(W + "line")
        out["lineRule"] = spacing.attrib.get(W + "lineRule")
        out["before"] = spacing.attrib.get(W + "before")
        out["after"] = spacing.attrib.get(W + "after")
    if ind is not None:
        out["firstLineChars"] = ind.attrib.get(W + "firstLineChars")
        out["leftChars"] = ind.attrib.get(W + "leftChars")
        out["hangingChars"] = ind.attrib.get(W + "hangingChars")
    if jc is not None:
        out["jc"] = jc.attrib.get(W + "val")
    return out


def show_first_run_font(node):
    run = node.find("w:r", NS)
    if run is None:
        return {}
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return {}
    rfonts = rpr.find("w:rFonts", NS)
    sz = rpr.find("w:sz", NS)
    bold = rpr.find("w:b", NS)
    return {
        "ascii": None if rfonts is None else rfonts.attrib.get(W + "ascii"),
        "eastAsia": None if rfonts is None else rfonts.attrib.get(W + "eastAsia"),
        "size": None if sz is None else sz.attrib.get(W + "val"),
        "bold": None if bold is None else bold.attrib.get(W + "val"),
    }


with ZipFile(DOCX) as zf:
    doc = ET.fromstring(zf.read("word/document.xml"))
    rels_xml = ET.fromstring(zf.read("word/_rels/document.xml.rels"))

rels = {rel.attrib.get("Id"): rel.attrib.get("Target") for rel in rels_xml}
body = doc.find("w:body", NS)
children = list(body)

print("== Figure area ==")
for idx in [171, 172, 173]:
    node = children[idx - 1]
    text = get_text(node)
    media = []
    for blip in node.findall(".//a:blip", NS):
        rid = blip.attrib.get("{%s}embed" % NS["r"])
        if rid:
            media.append(rels.get(rid, rid))
    print(idx, text, media, show_spacing(node))

print("\n== Sample headings ==")
for idx in [58, 60, 83, 84, 170, 174, 388]:
    node = children[idx - 1]
    print(idx, get_text(node), show_spacing(node), show_first_run_font(node))

print("\n== Table captions ==")
for idx in [134, 154, 157, 288, 342, 354]:
    node = children[idx - 1]
    print(idx, get_text(node), show_spacing(node), show_first_run_font(node))

print("\n== References ==")
for idx in range(388, min(396, len(children) + 1)):
    node = children[idx - 1]
    print(idx, get_text(node), show_spacing(node))

print("\n== Table 4-1 sample row ==")
tbl = children[344 - 1]
tpr = tbl.find("w:tblPr", NS)
if tpr is not None:
    borders = tpr.find("w:tblBorders", NS)
    if borders is not None:
        print("table borders", {child.tag.split('}')[1]: child.attrib.get(W + "val") for child in list(borders)})
for row_idx, row in enumerate(tbl.findall("w:tr", NS)[:3], 1):
    cells = [get_text(tc) for tc in row.findall("w:tc", NS)]
    print(row_idx, cells)
    if row_idx == 1:
        first_p = row.find("w:tc/w:p", NS)
        if first_p is not None:
            print("header font", show_first_run_font(first_p), show_spacing(first_p))
        first_tc = row.find("w:tc", NS)
        if first_tc is not None:
            tc_pr = first_tc.find("w:tcPr", NS)
            borders = None if tc_pr is None else tc_pr.find("w:tcBorders", NS)
            if borders is not None:
                print("header cell borders", {child.tag.split('}')[1]: child.attrib.get(W + "val") for child in list(borders)})
