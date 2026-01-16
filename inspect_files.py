
import pandas as pd
import os
from docx import Document

excel_path = r"d:\HACC_DEV\rapport mensuel\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx"
word_path = r"d:\HACC_DEV\rapport mensuel\REPORTING ADMINISTRATIF-FINANCIER-VENTES JUIN 2025.docx"

print("--- EXCEL FILE ANALYSIS ---")
try:
    # Read all sheets to see what's available
    xls = pd.ExcelFile(excel_path)
    print(f"Sheet names: {xls.sheet_names}")
    
    # Read the first sheet
    df = pd.read_excel(excel_path, sheet_name=0)
    print("\nFirst 5 rows of Sheet 1:")
    print(df.head())
    print("\nColumns:")
    print(df.columns.tolist())
except Exception as e:
    print(f"Error reading Excel: {e}")

print("\n--- WORD FILE ANALYSIS ---")
try:
    doc = Document(word_path)
    print("First 5 paragraphs:")
    for para in doc.paragraphs[:5]:
        print(para.text)
    
    print("\nTables found:", len(doc.tables))
    if len(doc.tables) > 0:
        print("First table, first row:")
        table = doc.tables[0]
        row = table.rows[0]
        print([cell.text for cell in row.cells])
except Exception as e:
    print(f"Error reading Word: {e}")
