import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def text_from_node(node):
    return "".join(t.text or "" for t in node.findall(".//w:t", NS)).strip()


def main():
    docx_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"E:\code\yolo_Rice_Pest\thesis_working.docx")
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(r"E:\code\yolo_Rice_Pest\thesis_extracted.txt")

    with ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")

    root = ET.fromstring(xml)
    body = root.find("w:body", NS)

    lines = []
    for idx, child in enumerate(body, 1):
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = text_from_node(child)
            if text:
                lines.append(f"P{idx}: {text}")
        elif tag == "tbl":
            lines.append(f"T{idx}:")
            for r, tr in enumerate(child.findall("w:tr", NS), 1):
                cells = []
                for tc in tr.findall("w:tc", NS):
                    cell_text = text_from_node(tc)
                    cells.append(cell_text)
                lines.append(f"  R{r}: {' | '.join(cells)}")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    print(str(out_path))


if __name__ == "__main__":
    main()
