# 📄 Excel → Word Report Generator

A pair of Python tools that convert an Excel (`.xlsx`) file into a formatted Word (`.docx`) report, complete with styled tables, proper borders, and clean typography.

> **Note:** Despite the `.js` file extensions, both scripts are Python files and must be run with Python.

---

## 🗂️ Tools Overview

| Tool | Script | Best For |
|------|--------|----------|
| **gen** | `gen.js` (Python) | Full vulnerability instances table reports — groups findings by issue & risk rating |
| **gen2** | `gen2.js` (Python) | Single-table instance lists — raw, flat table from Excel with ghost-row filtering |

---

## 📋 Prerequisites

- Python 3.8+
- Install dependencies:

```bash
pip install pandas openpyxl python-docx
```

> `gen.js` uses **pandas** for reading Excel.  
> `gen2.js` uses **openpyxl** directly (no pandas required).

---

## ⚡ Quick Start

### gen.js — All Vulnerability Instances Table Report Generator
```bash
py gen.js input.xlsx
py gen.js input.xlsx output.docx
```

### gen2.js — Instances Table Generator
```bash
py gen2.js input.xlsx
py gen2.js input.xlsx output.docx
```

**Output:** If no output path is given, both tools auto-generate `<input_name>_instances.docx` in the same directory.

---

## 🔍 Detailed Comparison

### Purpose & Output Structure

| Aspect | `gen.js` | `gen2.js` |
|--------|----------|----------|
| **Primary goal** | Full vulnerability report with multiple tables | Single instance table from raw Excel data |
| **Output structure** | One table **per unique vulnerability**, with heading and risk label | One flat table for the **entire sheet** |
| **Document title** | Adds `Vulnerability Report` heading | Adds `Instances:` heading |

---

### Excel Reading Strategy

| Aspect | `gen.js` | `gen2.js` |
|--------|----------|----------|
| **Library** | `pandas` (`read_excel`) | `openpyxl` (`load_workbook`) |
| **NaN / null handling** | `df.fillna('')` — fills all blanks | Iterates raw cells; skips `None` values per cell |
| **Ghost row filtering** | Not explicitly handled (pandas reads all rows) | ✅ Explicitly filters out rows where **all cells are blank** |
| **Header detection** | Excludes known meta-columns (`Security Issue`, `Risk Rating`, `Remark`, `References`) — uses remaining columns as dynamic instance columns | Reads the first row as headers; scans to find the **last non-empty header** to determine table width |

---

### Data Grouping & Logic

| Aspect | `gen.js` | `gen2.js` |
|--------|----------|----------|
| **Grouping** | Groups rows by `Security Issue` + `Risk Rating` | No grouping — all rows in a single flat table |
| **Column filtering** | Per-group: only renders columns that have **at least one non-empty value** | All columns up to the last named header are shown |
| **Dynamic column widths** | Always evenly distributes `16cm / num_cols` | Uses `COL_WIDTHS_CM = [4.5, 4.0, 5.0, 3.5]` if exactly 4 columns; falls back to equal distribution otherwise |
| **Sections per vulnerability** | ✅ Yes — Vulnerability Name (bold, 14pt) + Risk Rating label, then table | ❌ No — single table only |

---

### Column Deduplication (First Column)

Both tools suppress repeating values in the **first column** to improve readability:

```
# Example: "192.168.1.1" appears 3 times → shown only once at the top of the group
192.168.1.1    /admin     200
               /login     200
               /config    403
```

> `gen.js` only updates `last_col_0` when the value is non-empty.  
> `gen2.js` always updates `last_col_0`, so even an empty string resets the dedup tracker.

---

### Code Architecture

| Aspect | `gen.js` | `gen2.js` |
|--------|----------|----------|
| **Functions** | `generate_report()` — monolithic, all-in-one | `read_excel()` + `build_doc()` — separated concerns |
| **Entry point** | `generate_report(input, output)` called directly | `main()` orchestrates `read_excel()` → `build_doc()` → `save()` |
| **Debug output** | Prints only success/error | Prints row count: `Rows with data (incl. header): N` |

---

### Formatting & Styling

Both tools share identical styling logic:

| Property | Value |
|----------|-------|
| Font | Calibri (East Asia: SimSun) |
| Font size | 10pt (body), 11pt (heading) |
| Border color | `#BFBFBF` (light grey) |
| Line spacing | 1.0 |
| Page size | A4 (21 × 29.7 cm) |
| Margins | 2.54 cm on all sides |
| Header row borders | Top + Bottom: size 6 |
| Data row borders | Bottom: size 4 (size 6 on last row) |
| Left/Right borders | None (always 0) |

---

## 📊 When to Use Which

### Use `gen.js` when:
- Your Excel has **multiple vulnerabilities** listed together
- You need a **report-style document** with one section per vulnerability
- Your Excel includes `Security Issue` and `Risk Rating` columns
- Column presence should vary per vulnerability (empty columns auto-hidden per group)

### Use `gen2.js` when:
- You have a **single flat table** of instances (e.g., one vulnerability's affected URLs)
- You want **ghost/blank row filtering** to be handled automatically
- You prefer a **lightweight, dependency-minimal** approach (no pandas needed)
- You want **predictable column widths** when working with exactly 4 columns

---

## 📁 Expected Excel Structure

### For `gen.js`
Your Excel must have these columns (any order):

| Security Issue | Risk Rating | \[Column A\] | \[Column B\] | ... | Remark | References |
|---------------|-------------|--------------|--------------|-----|--------|------------|
| SQL Injection | High | 192.168.1.1 | /login | ... | | |
| XSS | Medium | 192.168.1.2 | /search | ... | | |

- `Security Issue` and `Risk Rating` are used for grouping (not shown in table)
- `Remark` and `References` are ignored
- All other columns become the instance table columns

### For `gen2.js`
Your Excel can be any simple table:

| No. | Affected Host | URL Path | HTTP Method |
|-----|--------------|----------|-------------|
| 1 | 192.168.1.1 | /admin | GET |
| 2 | 192.168.1.1 | /login | POST |

- First row = headers
- All subsequent rows = data
- Blank rows are automatically skipped

---

## 🛠️ Configuration

Both files share these constants at the top that you can edit:

```python
FONT_NAME     = "Calibri"       # Main font
FONT_EASTASIA = "SimSun"        # East Asian character font
FONT_SIZE     = Pt(10)          # Body text size
HEADING_SIZE  = Pt(11)          # Heading text size
BORDER_COLOR  = "BFBFBF"        # Table border colour (hex, no #)
```

**`gen2.js` only** — edit default column widths for 4-column tables:
```python
COL_WIDTHS_CM = [4.5, 4.0, 5.0, 3.5]
```

---

## 📦 File Reference

```
instance/
├── gen.js          # Vulnerability report generator (pandas-based)
├── gen2.js         # Instances table generator (openpyxl-based)
├── ins.xlsx        # Sample input Excel file
├── out.docx        # Sample output Word document
├── package.json    # Node dependencies (not used by Python scripts)
└── README.md       # This file
```

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| `ModuleNotFoundError: pandas` | Run `pip install pandas` |
| `ModuleNotFoundError: openpyxl` | Run `pip install openpyxl` |
| `ModuleNotFoundError: docx` | Run `pip install python-docx` |
| Output has blank/extra tables | Ensure `Security Issue` and `Risk Rating` cells are not empty in your Excel (for `gen.js`) |
| Columns missing in output | `gen.js` hides columns that are entirely empty for a vulnerability group — add data to those cells |
| Ghost rows appearing | Use `gen2.js` which explicitly filters empty rows, or clean your Excel file |
