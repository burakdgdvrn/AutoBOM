import pandas as pd
import json
from compare_bom import get_row_key, get_base_drawing

# Load Excel
excel_path = r"c:\Users\burak\Desktop\PDF okuma\E-26-RNS-01_BOM_Listesi (1).xlsx"
df = pd.read_excel(excel_path)
excel_rows = df.to_dict('records')

# Filter Excel for FG-502433-2001-P1401
filtered_excel = []
for r in excel_rows:
    dwg = str(r.get('Technical Drawing', ''))
    if get_base_drawing(dwg) == 'FG-502433-2001-P1401':
        filtered_excel.append(r)

print(f"Filtered Excel Rows for FG-502433-2001-P1401: {len(filtered_excel)}")

# Print first 5 Excel keys
for i, r in enumerate(filtered_excel[:5]):
    k = get_row_key(r, True)
    print(f"EXCEL ROW {i}: Assembly={r.get('Assembly')} | Sub={r.get('Sub Assembly')} | Item={r.get('Item Code')} -> KEY: {k}")

from pdf_to_excel import extract_bom_data_to_json

pdf_path = r"c:\Users\burak\Desktop\PDF okuma\PDFler\10-FG-502433-2001-P1401_1_1_Rev-0B.pdf"
json_data = extract_bom_data_to_json(pdf_path, "10-FG-502433-2001-P1401_1_1_Rev-0B")

print(f"\nAI Extracted Rows: {len(json_data)}")
for i, r in enumerate(json_data[:10]):
    if r.get('assembly') or r.get('item_code'):
        k = get_row_key(r, False)
        print(f"AI ROW {i}: Assembly={r.get('assembly')} | Sub={r.get('sub_assembly')} | Item={r.get('item_code')} -> KEY: {k}")

