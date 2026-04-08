from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
import re


IN_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v8.docx")
OUT_DOCX = Path(r"E:\code\yolo_Rice_Pest\thesis_modified_v9.docx")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{%s}" % NS["w"]


def ensure(parent, tag, first=False):
    node = parent.find(tag, NS)
    if node is None:
        node = ET.Element(W + tag.split(":")[1])
        if first and len(parent):
            parent.insert(0, node)
        else:
            parent.append(node)
    return node


def remove_all(parent, tag):
    for node in list(parent.findall(tag, NS)):
        parent.remove(node)


def get_text(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def clear_except_ppr(p):
    for child in list(p):
        if child.tag != W + "pPr":
            p.remove(child)


def set_spacing(ppr, *, line="500", before="0", after="0"):
    spacing = ensure(ppr, "w:spacing")
    spacing.set(W + "line", line)
    spacing.set(W + "lineRule", "exact")
    spacing.set(W + "before", before)
    spacing.set(W + "after", after)


def set_indent(ppr, *, first_line_chars=None, left_chars=None, hanging_chars=None):
    ind = ensure(ppr, "w:ind")
    for attr, value in (
        ("firstLineChars", first_line_chars),
        ("leftChars", left_chars),
        ("hangingChars", hanging_chars),
    ):
        if value is None:
            ind.attrib.pop(W + attr, None)
        else:
            ind.set(W + attr, str(value))
    if first_line_chars is not None:
        ind.set(W + "firstLine", "0" if str(first_line_chars) == "0" else "480")
    else:
        ind.attrib.pop(W + "firstLine", None)
    if left_chars is not None:
        ind.set(W + "left", "0")
    else:
        ind.attrib.pop(W + "left", None)
    if hanging_chars is not None:
        ind.set(W + "hanging", "0")
    else:
        ind.attrib.pop(W + "hanging", None)


def set_jc(ppr, val):
    jc = ensure(ppr, "w:jc")
    jc.set(W + "val", val)


def set_keep(ppr, *, keep_next=None, keep_lines=None):
    for tag, enabled in (("w:keepNext", keep_next), ("w:keepLines", keep_lines)):
        node = ppr.find(tag, NS)
        if enabled:
            if node is None:
                ET.SubElement(ppr, W + tag.split(":")[1])
                node = ppr.find(tag, NS)
            if node is not None:
                node.attrib.pop(W + "val", None)
        elif enabled is False and node is not None:
            ppr.remove(node)


def set_page_break_before(ppr, enabled):
    node = ppr.find("w:pageBreakBefore", NS)
    if enabled:
        if node is None:
            node = ET.SubElement(ppr, W + "pageBreakBefore")
        node.set(W + "val", "1")
    elif node is not None:
        ppr.remove(node)


def remove_num_pr(ppr):
    num = ppr.find("w:numPr", NS)
    if num is not None:
        ppr.remove(num)


def set_run_font(run, *, font="宋体", size="21", bold=False):
    rpr = ensure(run, "w:rPr", first=True)
    rfonts = ensure(rpr, "w:rFonts")
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(W + attr, font)
    sz = ensure(rpr, "w:sz")
    sz.set(W + "val", size)
    szcs = ensure(rpr, "w:szCs")
    szcs.set(W + "val", size)
    color = ensure(rpr, "w:color")
    color.set(W + "val", "000000")
    for tag in ("w:b", "w:bCs"):
        node = ensure(rpr, tag)
        node.set(W + "val", "1" if bold else "0")


def set_paragraph_text(p, text, *, font="宋体", size="21", bold=False):
    ppr = p.find("w:pPr", NS)
    clear_except_ppr(p)
    run = ET.Element(W + "r")
    t = ET.SubElement(run, W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    p.append(run)
    set_run_font(run, font=font, size=size, bold=bold)
    if ppr is not None and p[0] is not ppr:
        p.remove(ppr)
        p.insert(0, ppr)


def style_heading_center(p, text, *, bold=True):
    ppr = ensure(p, "w:pPr", first=True)
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        ppr.remove(ind)
    remove_num_pr(ppr)
    set_page_break_before(ppr, False)
    set_spacing(ppr, line="500", before="360", after="240")
    set_jc(ppr, "center")
    set_indent(ppr, first_line_chars=0, left_chars=None, hanging_chars=None)
    set_keep(ppr, keep_next=True, keep_lines=True)
    set_paragraph_text(p, text, font="黑体", size="36", bold=bold)


def style_ref_para(p, text):
    ppr = ensure(p, "w:pPr", first=True)
    remove_num_pr(ppr)
    set_page_break_before(ppr, False)
    set_spacing(ppr, line="500", before="0", after="0")
    set_jc(ppr, "left")
    set_indent(ppr, first_line_chars=200, left_chars=0, hanging_chars=None)
    set_keep(ppr, keep_next=False, keep_lines=True)
    set_paragraph_text(p, text, font="宋体", size="21", bold=False)
    keep = ensure(ppr, "w:keepLines")
    keep.attrib.pop(W + "val", None)


def style_ack_para(p):
    ppr = ensure(p, "w:pPr", first=True)
    remove_num_pr(ppr)
    set_page_break_before(ppr, False)
    set_spacing(ppr, line="500", before="0", after="0")
    set_jc(ppr, "both")
    set_indent(ppr, first_line_chars=200, left_chars=None, hanging_chars=None)
    set_keep(ppr, keep_next=False, keep_lines=False)
    for run in p.findall("w:r", NS):
        set_run_font(run, font="宋体", size="24", bold=False)


def make_footer_xml(mode):
    footer = ET.Element(W + "ftr")
    p = ET.SubElement(footer, W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    jc = ET.SubElement(ppr, W + "jc")
    jc.set(W + "val", "center")

    def add_run(text=None, field=None):
        r = ET.SubElement(p, W + "r")
        rpr = ET.SubElement(r, W + "rPr")
        rfonts = ET.SubElement(rpr, W + "rFonts")
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(W + attr, "宋体")
        sz = ET.SubElement(rpr, W + "sz")
        sz.set(W + "val", "21")
        szcs = ET.SubElement(rpr, W + "szCs")
        szcs.set(W + "val", "21")
        color = ET.SubElement(rpr, W + "color")
        color.set(W + "val", "000000")
        if field == "begin":
            fld = ET.SubElement(r, W + "fldChar")
            fld.set(W + "fldCharType", "begin")
        elif field == "instr":
            instr = ET.SubElement(r, W + "instrText")
            instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            instr.text = "PAGE"
        elif field == "sep":
            fld = ET.SubElement(r, W + "fldChar")
            fld.set(W + "fldCharType", "separate")
        elif field == "end":
            fld = ET.SubElement(r, W + "fldChar")
            fld.set(W + "fldCharType", "end")
        else:
            t = ET.SubElement(r, W + "t")
            if text and (text.startswith(" ") or text.endswith(" ")):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = text or ""

    if mode == "body":
        add_run("- ")
        add_run(field="begin")
        add_run(field="instr")
        add_run(field="sep")
        add_run("1")
        add_run(field="end")
        add_run(" -")
    elif mode == "roman":
        add_run(field="begin")
        add_run(field="instr")
        add_run(field="sep")
        add_run("I")
        add_run(field="end")
    return ET.tostring(footer, encoding="utf-8", xml_declaration=True)


def make_blank_header_xml():
    hdr = ET.Element(W + "hdr")
    p = ET.SubElement(hdr, W + "p")
    ppr = ET.SubElement(p, W + "pPr")
    jc = ET.SubElement(ppr, W + "jc")
    jc.set(W + "val", "center")
    return ET.tostring(hdr, encoding="utf-8", xml_declaration=True)


def set_sectpr_refs(sect, *, default_header=None, even_header=None, default_footer=None, even_footer=None, first_footer=None):
    remove_all(sect, "w:headerReference")
    remove_all(sect, "w:footerReference")
    insert_at = 0
    for typ, rid, local in (
        ("default", default_header, "headerReference"),
        ("even", even_header, "headerReference"),
        ("first", None, "headerReference"),
        ("first", first_footer, "footerReference"),
        ("default", default_footer, "footerReference"),
        ("even", even_footer, "footerReference"),
    ):
        if rid is None:
            continue
        node = ET.Element(W + local)
        node.set(W + "type", typ)
        node.set("{%s}id" % NS["r"], rid)
        sect.insert(insert_at, node)
        insert_at += 1


def set_pgnum(sect, *, fmt, start=None):
    pg = ensure(sect, "w:pgNumType")
    pg.set(W + "fmt", fmt)
    if start is None:
        pg.attrib.pop(W + "start", None)
    else:
        pg.set(W + "start", str(start))


def set_sect_type(sect, val=None):
    node = sect.find("w:type", NS)
    if val is None:
        if node is not None:
            sect.remove(node)
    else:
        if node is None:
            node = ET.SubElement(sect, W + "type")
        node.set(W + "val", val)


def add_or_replace_sectpr(p, new_sect):
    ppr = ensure(p, "w:pPr", first=True)
    old = ppr.find("w:sectPr", NS)
    if old is not None:
        ppr.remove(old)
    ppr.append(new_sect)


def copy_tail_sect(children):
    tail = children[418 - 1]
    if tail.tag == W + "sectPr":
        return deepcopy(tail)
    ppr = ensure(tail, "w:pPr", first=True)
    return deepcopy(ensure(ppr, "w:sectPr"))


def clean_ref_text(text, idx):
    cleaned = re.sub(r"^\[\d+\](\[\d+\])?\s*", "", text).strip()
    return f"[{idx}] {cleaned}"


with ZipFile(IN_DOCX) as zf:
    files = {name: zf.read(name) for name in zf.namelist()}

doc = ET.fromstring(files["word/document.xml"])
body = doc.find("w:body", NS)
children = list(body)

# Rewrite header/footer parts to target formats.
files["word/footer1.xml"] = make_footer_xml("body")
files["word/footer2.xml"] = make_footer_xml("body")
files["word/footer3.xml"] = make_footer_xml("roman")
files["word/footer4.xml"] = make_footer_xml("body")
files["word/footer5.xml"] = make_footer_xml("body")
files["word/header5.xml"] = make_blank_header_xml()
files["word/header6.xml"] = make_blank_header_xml()

# TOC section footer/header and roman numbering.
toc_sect_p = children[53 - 1]
toc_ppr = ensure(toc_sect_p, "w:pPr", first=True)
toc_sect = ensure(toc_ppr, "w:sectPr")
set_sectpr_refs(toc_sect, default_header="rId3", even_header="rId4", default_footer="rId9", even_footer="rId9")
set_pgnum(toc_sect, fmt="upperRoman", start=1)
set_sect_type(toc_sect, "oddPage")

# Body section break before references on paragraph 387.
body_break_p = children[387 - 1]
clear_except_ppr(body_break_p)
body_break_ppr = ensure(body_break_p, "w:pPr", first=True)
remove_num_pr(body_break_ppr)
set_spacing(body_break_ppr, line="500", before="0", after="0")
set_jc(body_break_ppr, "left")
set_indent(body_break_ppr, first_line_chars=0, left_chars=None, hanging_chars=None)
set_keep(body_break_ppr, keep_next=False, keep_lines=False)
body_sect = copy_tail_sect(children)
set_sectpr_refs(body_sect, default_header="rId5", even_header="rId6", default_footer="rId7", even_footer="rId8")
set_pgnum(body_sect, fmt="decimal", start=1)
set_sect_type(body_sect, "nextPage")
add_or_replace_sectpr(body_break_p, body_sect)

# Reference heading and entries.
style_heading_center(children[388 - 1], "参考文献", bold=True)
for ref_idx, body_idx in enumerate(range(389, 409), 1):
    text = get_text(children[body_idx - 1])
    style_ref_para(children[body_idx - 1], clean_ref_text(text, ref_idx))

# Reference section ends on paragraph 408 and ack starts next page.
ref_end_p = children[408 - 1]
ref_end_ppr = ensure(ref_end_p, "w:pPr", first=True)
ref_sect = copy_tail_sect(children)
set_sectpr_refs(ref_sect, default_header="rId10", even_header="rId11", default_footer="rId12", even_footer="rId13")
set_pgnum(ref_sect, fmt="decimal", start=None)
set_sect_type(ref_sect, "nextPage")
add_or_replace_sectpr(ref_end_p, ref_sect)
style_ref_para(ref_end_p, clean_ref_text(get_text(ref_end_p), 20))

# Acknowledgement section formatting.
style_heading_center(children[409 - 1], "致谢", bold=True)
for body_idx in range(410, 417):
    style_ack_para(children[body_idx - 1])

# Final section for acknowledgement: no header, body-style footer.
final_sect = children[418 - 1]
set_sectpr_refs(final_sect, default_header="rId10", even_header="rId11", default_footer="rId12", even_footer="rId13")
set_pgnum(final_sect, fmt="decimal", start=None)
set_sect_type(final_sect, None)

files["word/document.xml"] = ET.tostring(doc, encoding="utf-8", xml_declaration=True)

with ZipFile(OUT_DOCX, "w", ZIP_DEFLATED) as zf:
    for name, data in files.items():
        zf.writestr(name, data)

print(str(OUT_DOCX))
