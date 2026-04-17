import sys
import os
import pandas as pd
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

# ── Main builder ───────────────────────────────────────────────────────────────
def generate_report(input_file, output_file):
    try:
        # 1. Read Excel and standardize missing values
        df = pd.read_excel(input_file)
        df = df.fillna('')

        # 2. DYNAMIC COLUMN DETECTION
        # Ignore these specific meta-columns from the Excel sheet. 
        # ANY other column found will be automatically used for the Instances table.
        ignore_headers = ['Security Issue', 'Risk Rating', 'Remark', 'References']
        
        # Extract headers directly from the Excel file (ignores blank "Unnamed" columns)
        base_columns = [col for col in df.columns if col not in ignore_headers and not str(col).startswith('Unnamed')]

        if not base_columns:
            raise ValueError("Could not find any valid instance columns in the Excel file.")

        # 3. Setup Word Document strictly to spec
        doc = Document()
        section = doc.sections[0]
        section.page_width    = Cm(21)
        section.page_height   = Cm(29.7)
        section.left_margin   = Cm(2.54)
        section.right_margin  = Cm(2.54)
        section.top_margin    = Cm(2.54)
        section.bottom_margin = Cm(2.54)

        doc.add_heading('Vulnerability Report', 0)

        # 4. Group logically by Issue and Severity
        grouped = df.groupby(['Security Issue', 'Risk Rating'], dropna=False)

        for (issue, risk), group in grouped:
            if str(issue).strip() == '' and str(risk).strip() == '':
                continue

            # ── DYNAMIC COLUMN FILTERING ──
            cols_to_keep = []
            for col in base_columns:
                if any(str(val).strip() != '' for val in group[col]):
                    cols_to_keep.append(col)
            
            # Fallback if a table has absolutely no data
            if not cols_to_keep:
                cols_to_keep = [base_columns[0]]

            num_cols = len(cols_to_keep)
            widths_cm = [16.0 / num_cols] * num_cols

            # ── Header: Vulnerability Name ──
            heading = doc.add_paragraph()
            heading.paragraph_format.space_after = Pt(0)
            heading.paragraph_format.line_spacing = 1.0 
            
            run = heading.add_run(str(issue))
            run.font.name  = FONT_NAME
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_EASTASIA)
            run.font.size  = Pt(14)
            run.font.bold  = True
            run.font.color.rgb = RGBColor(0, 0, 0)

            # ── Subheader: Risk Rating ──
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6) 
            
            r1 = p.add_run('Risk Rating: ')
            r1.font.name = FONT_NAME
            r1.font.bold = True
            r1.font.size = FONT_SIZE
            
            r2 = p.add_run(str(risk))
            r2.font.name = FONT_NAME
            r2.font.size = FONT_SIZE

            # ── Custom Formatted Table ──
            table = doc.add_table(rows=0, cols=num_cols)
            table.style = "Normal Table"

            # Render Table Headers natively from Excel names
            hdr_row = table.add_row()
            for i, label in enumerate(cols_to_keep):
                cell = hdr_row.cells[i]
                cell.width = Cm(widths_cm[i])
                set_cell_text(cell, label, bold=True)
                clear_cell_shading(cell)
                set_cell_borders(cell, top=6, bottom=6, left=0, right=0)

            # Render Data Rows
            data_rows = group[cols_to_keep].to_dict('records')
            last_col_0 = None
            
            for row_idx, row_data in enumerate(data_rows):
                is_last = (row_idx == len(data_rows) - 1)
                tr = table.add_row()

                for i, col_name in enumerate(cols_to_keep):
                    val = str(row_data[col_name]).strip()

                    # Hides repeating values in the FIRST column
                    if i == 0:
                        display_val = "" if val == last_col_0 else val
                        if val != "":
                            last_col_0 = val
                    else:
                        display_val = val

                    cell = tr.cells[i]
                    cell.width = Cm(widths_cm[i])
                    set_cell_text(cell, display_val, bold=False)
                    clear_cell_shading(cell)
                    
                    bottom_size = 6 if is_last else 4
                    set_cell_borders(cell, top=0, bottom=bottom_size, left=0, right=0)

            doc.add_paragraph()

        doc.save(output_file)
        print(f"✅ Success! Report generated with dynamic headers at: {output_file}")

    except Exception as e:
        print(f"❌ An error occurred: {e}")

# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: py gen.py input.xlsx [output.docx]")
        sys.exit(1)

    input_path  = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base        = os.path.splitext(input_path)[0]
        output_path = base + "_instances.docx"

    print(f"Reading:  {input_path}")
    generate_report(input_path, output_path)