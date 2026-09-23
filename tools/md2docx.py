#!/usr/bin/env python3
"""Chuyển đổi báo cáo giải trình mua sắm (Markdown) sang bản Word (.docx).

Định dạng: khổ A4, font Times New Roman 13 pt (chuẩn văn bản hành chính);
bảng >= 8 cột (bảng tổng hợp 11 cột) tự động đặt trong section khổ ngang.
Cách dùng:  python3 tools/md2docx.py [file.md] [file.docx]
"""
import re
import sys

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SRC = sys.argv[1] if len(sys.argv) > 1 else "bao-cao-giai-trinh-mua-sam-phan-bo-thiet-bi.md"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Bao-cao-giai-trinh-mua-sam-phan-bo-thiet-bi.docx"

INLINE = re.compile(r"(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`)")
SEP_CELL = re.compile(r"^:?-{2,}:?$")


def add_inline(par, text, base_bold=False, size=None):
    """Thêm text vào paragraph, hỗ trợ **bold**, *italic*, `code`, <br>."""
    for tok in INLINE.split(text.replace("<br>", "\n")):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            run = par.add_run(tok[2:-2])
            run.bold = True
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            run = par.add_run(tok[1:-1])
            run.italic = True
        elif tok.startswith("`") and tok.endswith("`"):
            run = par.add_run(tok[1:-1])
            run.font.name = "Consolas"
        else:
            run = par.add_run(tok)
            if base_bold:
                run.bold = True
        if size is not None:
            run.font.size = Pt(size)


def set_cell_margins_small(table):
    tbl_pr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("left", 60), ("right", 60), ("top", 30), ("bottom", 30)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)


def repeat_header(table):
    tr = table.rows[0]._tr
    tr_pr = tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def add_table(doc, rows, landscape):
    n_cols = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Table Grid"
    table.autofit = True
    set_cell_margins_small(table)
    repeat_header(table)
    font_size = 8.5 if n_cols >= 8 else (10 if n_cols >= 5 else 11)
    for i, row in enumerate(rows):
        for j, cell_text in enumerate(row):
            cell = table.cell(i, j)
            par = cell.paragraphs[0]
            par.paragraph_format.space_after = Pt(2)
            par.paragraph_format.space_before = Pt(2)
            add_inline(par, cell_text, base_bold=(i == 0), size=font_size)
    return table


def new_section(doc, landscape):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    return sec


def main():
    with open(SRC, encoding="utf-8") as fh:
        lines = fh.read().splitlines()

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(13)
    for name, size in (("Title", 18), ("Heading 1", 15), ("Heading 2", 13.5), ("Heading 3", 13)):
        st = doc.styles[name]
        st.font.name = "Times New Roman"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
    first = doc.sections[0]
    first.orientation = WD_ORIENT.PORTRAIT
    first.page_width, first.page_height = Cm(21.0), Cm(29.7)  # A4
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(1.8)
        sec.top_margin = sec.bottom_margin = Cm(1.8)

    in_landscape = False
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        # --- bảng ---
        if stripped.startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(SEP_CELL.match(c) for c in cells):
                    block.append(cells)
                i += 1
            n_cols = max(len(r) for r in block)
            block = [r + [""] * (n_cols - len(r)) for r in block]
            need_landscape = n_cols >= 8
            if need_landscape != in_landscape:
                new_section(doc, need_landscape)
                in_landscape = need_landscape
            add_table(doc, block, in_landscape)
            doc.add_paragraph()
            continue

        # --- heading ---
        if stripped.startswith("# "):
            par = doc.add_heading(level=0)
            add_inline(par, stripped[2:])
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif stripped.startswith("## "):
            par = doc.add_heading(level=1)
            add_inline(par, stripped[3:])
        elif stripped.startswith("### "):
            par = doc.add_heading(level=2)
            add_inline(par, stripped[4:])
        elif stripped == "---":
            pass
        elif stripped.startswith("> "):
            par = doc.add_paragraph()
            add_inline(par, stripped[2:])
            for run in par.runs:
                run.italic = True
                run.font.size = Pt(12)
        elif stripped.startswith("- "):
            par = doc.add_paragraph(style="List Bullet")
            add_inline(par, stripped[2:])
        elif re.match(r"^ {2,}- ", line):
            par = doc.add_paragraph(style="List Bullet 2")
            add_inline(par, line.strip()[2:])
        elif re.match(r"^\d+\. ", stripped):
            par = doc.add_paragraph(style="List Number")
            add_inline(par, re.sub(r"^\d+\. ", "", stripped))
        elif stripped:
            par = doc.add_paragraph()
            add_inline(par, stripped)
        i += 1

    doc.save(OUT)
    print(f"Đã tạo {OUT}: {len(doc.paragraphs)} đoạn, {len(doc.tables)} bảng, {len(doc.sections)} section")


if __name__ == "__main__":
    main()
