#!/usr/bin/env python3
"""
Excel → DOCX Report Generator
Usage: python3 excel_to_docx.py input.xlsx output.docx

The Excel file must have two columns:
  - Misconfiguration   : The finding title/heading
  - CSTP Justification : The explanation text
"""

import sys
import json
import subprocess
import os
import tempfile
import pandas as pd

def main():
    if len(sys.argv) < 3:
        print("Usage: python3 excel_to_docx.py input.xlsx output.docx")
        sys.exit(1)

    input_xlsx = sys.argv[1]
    output_docx = sys.argv[2]

    if not os.path.exists(input_xlsx):
        print(f"Error: File not found: {input_xlsx}")
        sys.exit(1)

    df = pd.read_excel(input_xlsx)

    required = {'Misconfiguration', 'CSTP Justification'}
    missing = required - set(df.columns)
    if missing:
        print(f"Error: Missing columns in Excel: {missing}")
        print(f"Found columns: {df.columns.tolist()}")
        sys.exit(1)

    df = df.fillna('')
    rows = df[['Misconfiguration', 'CSTP Justification']].to_dict(orient='records')

    # Write data to a temp JSON file to avoid Windows command line length limit
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8')
    json.dump(rows, tmp, ensure_ascii=False)
    tmp.close()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    js_script = os.path.join(script_dir, 'format.js')

    try:
        result = subprocess.run(
            ['node', js_script, tmp.name, output_docx],
            capture_output=True, text=True
        )
    finally:
        # Clean up temp file if node didn't already
        if os.path.exists(tmp.name):
            os.unlink(tmp.name)

    if result.returncode != 0:
        print("Error generating DOCX:")
        print(result.stderr)
        sys.exit(1)

    print(result.stdout.strip())
    print(f"Report saved to: {output_docx}")

if __name__ == '__main__':
    main()