"""
Excel to Word Converter — Instances Table
==========================================
Reads an Excel file and dynamically uses its headers.
Filters out blank "ghost" rows automatically.
"""

import sys
import os
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Config ─────────────────────────────────────────────────────────────────────
FONT_NAME      = "Calibri"
FONT_EASTASIA  = "SimSun"
FONT_SIZE      = Pt(10)
HEADING_SIZE   = Pt(11)
BORDER_COLOR   = "BFBFBF"   # hex without #

# Default column widths in cm if you have exactly 4 columns.
# (If you have more/less, the script will distribute space automatically)
COL_WIDTHS_CM  = [4.5, 4.0, 5.0, 3.5]


# ── Border helpers ─────────────────────────────────────────────────────────────
def _make_border_element(border_type, size, color):
    el = OxmlElement(f"w:{border_type}")
    el.set(qn("w:val"),   "single" if size else "none")
    el.set(qn("w:sz"),    str(size))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color if size else "FFFFFF")
    return el


def set_cell_borders(cell, top=0, bottom=0, left=0, right=0):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()

    existing = tcPr.find(qn("w:tcBorders"))
    if existing is not None:
        tcPr.remove(existing)

    tcBorders = OxmlElement("w:tcBorders")
    tcBorders.append(_make_border_element("top",    top,    BORDER_COLOR))
    tcBorders.append(_make_border_element("left",   left,   BORDER_COLOR))
    tcBorders.append(_make_border_element("bottom", bottom, BORDER_COLOR))
    tcBorders.append(_make_border_element("right",  right,  BORDER_COLOR))
    tcPr.append(tcBorders)


def clear_cell_shading(cell):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = tcPr.find(qn("w:shd"))
    if shd is not None:
        tcPr.remove(shd)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "auto")
    tcPr.append(shd)


# ── Cell content ───────────────────────────────────────────────────────────────
def set_cell_text(cell, text, bold=False):
    """Write text into a cell with formatting and 1.0 line spacing."""
    cell.text = ""
    para = cell.paragraphs[0]
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    para.paragraph_format.line_spacing = 1.0 
    
    run = para.add_run(text)
    run.font.name  = FONT_NAME
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_EASTASIA)
    run.font.size  = FONT_SIZE
    run.font.bold  = bold
    run.font.color.rgb = RGBColor(0, 0, 0)


# ── Read Excel ─────────────────────────────────────────────────────────────────
def read_excel(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = []
    
    for row in ws.iter_rows(values_only=True):
        cleaned_row = [str(c).strip() if c is not None else "" for c in row]
        
        # FIX: Only append the row if it contains actual data.
        # This completely ignores "ghost" blank rows in Excel.
        if any(cleaned_row):
            rows.append(cleaned_row)
            
    return rows


# ── Main builder ───────────────────────────────────────────────────────────────
def build_doc(rows):
    if not rows:
        raise ValueError("Excel file is empty or contains no readable data.")

    header_row, *data_rows = rows

    # FIX: Dynamically read headers from the Excel file
    # We find the last column that actually has a name to determine the table size
    last_non_empty = -1
    for i, h in enumerate(header_row):
        if h: 
            last_non_empty = i
            
    if last_non_empty == -1:
        raise ValueError("Excel file has no headers in the first row.")
        
    # Set the display headers exactly as they appear in Excel
    display_headers = header_row[:last_non_empty + 1]
    num_cols = len(display_headers)

    # If the number of columns matches our config, use standard widths.
    # Otherwise, distribute the A4 page width (approx 16cm) evenly.
    if num_cols == len(COL_WIDTHS_CM):
        widths_cm = COL_WIDTHS_CM
    else:
        widths_cm = [16.0 / num_cols] * num_cols

    doc = Document()

    # ── Page setup: A4 with 1-inch margins ──
    section = doc.sections[0]
    section.page_width    = Cm(21)
    section.page_height   = Cm(29.7)
    section.left_margin   = Cm(2.54)
    section.right_margin  = Cm(2.54)
    section.top_margin    = Cm(2.54)
    section.bottom_margin = Cm(2.54)

    # ── "Instances:" heading ──
    heading = doc.add_paragraph()
    heading.paragraph_format.space_after = Pt(0)
    heading.paragraph_format.line_spacing = 1.0 
    
    run = heading.add_run("Instances:")
    run.font.name  = FONT_NAME
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_EASTASIA)
    run.font.size  = HEADING_SIZE
    run.font.bold  = True
    run.font.color.rgb = RGBColor(0, 0, 0)

    # ── Table ──
    table = doc.add_table(rows=0, cols=num_cols)
    table.style = "Normal Table"

    # Set column widths
    for i, width_cm in enumerate(widths_cm):
        for cell in table.columns[i].cells:
            cell.width = Cm(width_cm)

    # ── Header row ──
    hdr_row = table.add_row()
    for i, label in enumerate(display_headers):
        cell = hdr_row.cells[i]
        cell.width = Cm(widths_cm[i])
        set_cell_text(cell, label, bold=True)
        clear_cell_shading(cell)
        set_cell_borders(cell, top=6, bottom=6, left=0, right=0)

    # ── Data rows ──
    last_col_0 = None
    for row_idx, row in enumerate(data_rows):
        is_last = (row_idx == len(data_rows) - 1)
        tr = table.add_row()
        
        # Ensure row has enough columns (pads with empty strings if missing)
        padded_row = row + [""] * (num_cols - len(row))

        for i in range(num_cols):
            val = padded_row[i]
            
            # Keep the logic that hides repeating values in the FIRST column
            if i == 0:
                display_val = "" if val == last_col_0 else val
                last_col_0 = val
            else:
                display_val = val
                
            cell = tr.cells[i]
            cell.width = Cm(widths_cm[i])
            set_cell_text(cell, display_val, bold=False)
            clear_cell_shading(cell)
            bottom_size = 6 if is_last else 4
            set_cell_borders(cell, top=0, bottom=bottom_size, left=0, right=0)

    return doc


# ── Entry point ────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_to_word_instances.py input.xlsx [output.docx]")
        sys.exit(1)

    input_path  = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base        = os.path.splitext(input_path)[0]
        output_path = base + "_instances.docx"

    print(f"Reading:  {input_path}")
    rows = read_excel(input_path)
    print(f"Rows with data (incl. header): {len(rows)}")

    doc = build_doc(rows)
    doc.save(output_path)
    print(f"Written:  {output_path}")


if __name__ == "__main__":
    main()