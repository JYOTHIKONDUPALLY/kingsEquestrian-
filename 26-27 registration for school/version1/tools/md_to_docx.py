#!/usr/bin/env python3
"""Convert TRAINER_STABLE_MANAGER_GUIDE.md to Word (.docx)."""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
MD_FILE = ROOT / "TRAINER_STABLE_MANAGER_GUIDE.md"
OUT_FILE = ROOT / "TRAINER_STABLE_MANAGER_GUIDE.docx"

BRAND = RGBColor(0x1F, 0x4E, 0x3D)


def strip_md_bold(text):
    return re.sub(r"\*\*(.+?)\*\*", r"\1", text)


def add_rich_paragraph(doc, text, style=None):
    """Paragraph with **bold** segments."""
    p = doc.add_paragraph(style=style)
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = p.add_run(part[2:-2])
            run.bold = True
        elif part:
            p.add_run(part)
    return p


def parse_table_lines(lines):
    rows = []
    for line in lines:
        line = line.strip()
        if not line.startswith("|"):
            continue
        if re.match(r"^\|[\s\-:|]+\|$", line):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
    return rows


def add_table(doc, rows):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncols):
            cell_text = row[j] if j < len(row) else ""
            cell = table.rows[i].cells[j]
            cell.text = strip_md_bold(cell_text)
            if i == 0:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
    doc.add_paragraph()


def convert(md_path, out_path):
    text = md_path.read_text(encoding="utf-8")
    lines = text.splitlines()
    doc = Document()

    # Default font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    i = 0
    in_ul = False
    ul_items = []

    def flush_ul():
        nonlocal in_ul, ul_items
        for item in ul_items:
            add_rich_paragraph(doc, item, style="List Bullet")
        ul_items = []
        in_ul = False

    while i < len(lines):
        line = lines[i]

        # Table block
        if line.strip().startswith("|"):
            flush_ul()
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i])
                i += 1
            add_table(doc, parse_table_lines(tbl_lines))
            continue

        stripped = line.strip()

        if stripped == "---":
            flush_ul()
            doc.add_paragraph()
            i += 1
            continue

        if stripped.startswith("# "):
            flush_ul()
            doc.add_heading(strip_md_bold(stripped[2:]), level=0)
            i += 1
            continue
        if stripped.startswith("## "):
            flush_ul()
            doc.add_heading(strip_md_bold(stripped[3:]), level=1)
            i += 1
            continue
        if stripped.startswith("### "):
            flush_ul()
            doc.add_heading(strip_md_bold(stripped[4:]), level=2)
            i += 1
            continue

        if re.match(r"^\d+\.\s", stripped):
            flush_ul()
            add_rich_paragraph(doc, re.sub(r"^\d+\.\s", "", stripped), style="List Number")
            i += 1
            continue

        if stripped.startswith("- [ ] "):
            flush_ul()
            add_rich_paragraph(doc, "☐ " + stripped[6:], style="List Bullet")
            i += 1
            continue

        if stripped.startswith("- "):
            if not in_ul:
                in_ul = True
            ul_items.append(stripped[2:])
            i += 1
            continue

        if not stripped:
            flush_ul()
            i += 1
            continue

        flush_ul()
        if stripped.startswith("*") and stripped.endswith("*") and not stripped.startswith("**"):
            p = doc.add_paragraph(stripped.strip("*"))
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for run in p.runs:
                run.italic = True
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        else:
            add_rich_paragraph(doc, stripped)
        i += 1

    flush_ul()

    # Footer note on title page area - document properties
    core = doc.core_properties
    core.title = "Trainer & Stable Manager Guide"
    core.subject = "Kings Equestrian Stable Management App"
    core.author = "Kings Equestrian Foundation"

    doc.save(out_path)
    print(f"Created: {out_path}")


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else MD_FILE
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else OUT_FILE
    convert(src, dst)
