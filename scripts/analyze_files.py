import pandas as pd
from bs4 import BeautifulSoup
import json
import sys

# Analysis of Excel file
excel_file = "c:\\Users\\burak\\Desktop\\PDF okuma\\E-26-RNS-01_BOM_Listesi (1).xlsx"
print("--- EXCEL FILE ANALYSIS ---")
try:
    xl = pd.ExcelFile(excel_file)
    print(f"Sheet names: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        print(f"\nSheet: {sheet}")
        df = xl.parse(sheet, nrows=5)
        print("Columns:")
        print(list(df.columns))
        print("First 3 rows:")
        print(df.head(3).to_string())
except Exception as e:
    print(f"Error reading Excel: {e}")

# Analysis of HTML file
html_file = "c:\\Users\\burak\\Desktop\\PDF okuma\\src\\ERP Sistemi.html"
print("\n--- HTML FILE ANALYSIS ---")
try:
    with open(html_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')
    
    inputs = soup.find_all(['input', 'select', 'textarea'])
    print(f"Total input fields found: {len(inputs)}")
    
    print("\nInput fields details (first 20):")
    for i, inp in enumerate(inputs[:20]):
        tag = inp.name
        inp_type = inp.get('type', 'text') if tag == 'input' else tag
        inp_id = inp.get('id', '')
        inp_name = inp.get('name', '')
        # Try to find associated label
        label_text = ''
        if inp_id:
            label = soup.find('label', attrs={'for': inp_id})
            if label:
                label_text = label.get_text(strip=True)
        if not label_text:
            # Maybe inside a label
            parent_label = inp.find_parent('label')
            if parent_label:
                label_text = parent_label.get_text(strip=True)
                
        print(f"{i+1}. Tag: {tag}, Type: {inp_type}, ID: {inp_id}, Name: {inp_name}, Label: {label_text}")

    # Check for tables
    tables = soup.find_all('table')
    print(f"\nTotal tables found: {len(tables)}")
    for i, table in enumerate(tables):
        headers = [th.get_text(strip=True) for th in table.find_all('th')]
        print(f"Table {i+1} headers: {headers}")

except Exception as e:
    print(f"Error reading HTML: {e}")
