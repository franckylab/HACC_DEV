
import pandas as pd
import os
import json
from docx import Document
import glob
from datetime import datetime
import re
import win32com.client as win32
import pythoncom

DATA_DIR = r"d:\HACC_DEV\rapport mensuel"
OUTPUT_FILE = r"d:\HACC_DEV\sales_data.json"

all_data = []

def convert_doc_to_docx(doc_path):
    try:
        pythoncom.CoInitialize()
        try:
            word = win32.DispatchEx("Word.Application")
        except:
            word = win32.Dispatch("Word.Application")
            
        word.Visible = False
    except Exception as e:
        print(f"  Cannot start Word for {doc_path}: {e}")
        return None

    docx_path = doc_path + "x"
    try:
        wb = word.Documents.Open(doc_path)
        wb.SaveAs2(docx_path, FileFormat=16)
        wb.Close()
        return docx_path
    except Exception as e:
        print(f"Error converting {doc_path}: {e}")
        return None
    finally:
        try:
            word.Quit()
        except:
            pass

def clean_amount(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        # Remove spaces, currency symbols, handle commas
        val = val.replace(' ', '').replace('FCFA', '').replace('F', '').strip()
        val = val.replace(',', '.') # Assume comma is decimal if mixed, but check format
        # Actually, "26 000" -> 26000. "20 000" -> 20000.
        # Sometimes comma is used for decimals, sometimes not. 
        # Let's assume standard int/float parsing after space removal.
        try:
            return float(val)
        except:
            return 0.0
    return 0.0

def clean_qty(val):
    if pd.isna(val):
        return 0.0
    try:
        return float(clean_amount(val))
    except:
        return 0.0

def parse_date(val):
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        # Try formats: DD/MM/YYYY, YYYY-MM-DD
        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
            try:
                return datetime.strptime(val.strip(), fmt)
            except:
                continue
    return None

def process_excel(filepath):
    print(f"Processing Excel: {os.path.basename(filepath)}")
    try:
        # Read Excel - Try 'VENTES' sheet, else first sheet
        xl = pd.ExcelFile(filepath)
        sheet_name = 'VENTES' if 'VENTES' in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        # Normalize columns
        df.columns = [str(c).strip() for c in df.columns]
        
        # Identify key columns
        date_col = next((c for c in df.columns if 'Date' in c), None)
        client_col = next((c for c in df.columns if 'Client' in c), None)
        qty_col = next((c for c in df.columns if 'Quantité' in c or 'Qte' in c), None)
        amount_col = next((c for c in df.columns if 'Montant' in c), None)
        
        if not (date_col and amount_col):
            print(f"  Skipping {filepath}: Missing Date or Amount columns")
            return

        for _, row in df.iterrows():
            d = parse_date(row[date_col])
            if not d: continue
            
            amt = clean_amount(row[amount_col])
            qty = clean_qty(row[qty_col]) if qty_col else 0
            client = str(row[client_col]).strip() if client_col and not pd.isna(row[client_col]) else "Unknown"
            
            all_data.append({
                "date": d.strftime("%Y-%m-%d"),
                "month": d.strftime("%B"),
                "year": d.year,
                "client": client,
                "quantity": qty,
                "revenue": amt,
                "source": os.path.basename(filepath)
            })
            
    except Exception as e:
        print(f"  Error processing Excel {filepath}: {e}")

def process_docx(filepath):
    print(f"Processing Word: {os.path.basename(filepath)}")
    try:
        doc = Document(filepath)
        target_table = None
        header_map = {}
        
        # Find the best sales table
        best_table = None
        best_score = 0
        best_map = {}

        for table in doc.tables:
            if not table.rows: continue
            first_row = [cell.text.strip().lower() for cell in table.rows[0].cells]
            
            # Score this table
            current_map = {}
            score = 0
            
            for idx, h in enumerate(first_row):
                if 'date' in h: 
                    current_map['date'] = idx
                    score += 1
                elif 'client' in h and 'nom' in h: 
                    current_map['client'] = idx
                    score += 2
                elif 'qte' in h or 'qté' in h: 
                    current_map['qty'] = idx
                    score += 2
                elif 'montant' in h: 
                    current_map['amount'] = idx
                    score += 1
                elif 'article' in h or 'code' in h or 'désignation' in h: 
                     if 'article' in h: current_map['product'] = idx
                     score += 3
            
            # Secondary checks
            if 'client' not in current_map:
                for idx, h in enumerate(first_row):
                     if 'client' in h: current_map['client'] = idx; score += 1; break
            
            if 'product' not in current_map:
                for idx, h in enumerate(first_row):
                     if 'article' in h or 'produit' in h or 'designation' in h or 'désignation' in h: 
                        current_map['product'] = idx; score += 2; break
            
            # Must have date and amount to even be considered
            if 'date' in current_map and 'amount' in current_map:
                if score > best_score:
                    best_score = score
                    best_table = table
                    best_map = current_map

        if not best_table:
            print(f"  Skipping {filepath}: No valid sales table found")
            return

        target_table = best_table
        header_map = best_map
        print(f"  Selected table with score {best_score}. Mapped headers: {header_map}")
        
        # Extract data
        count = 0
        for i, row in enumerate(target_table.rows[1:], start=1):
            cells = row.cells
            try:
                date_str = cells[header_map['date']].text.strip()
                # Skip empty date rows
                if not date_str: continue

                d = parse_date(date_str)
                if not d: 
                    # print(f"    Skipping row {i}: Invalid date '{date_str}'")
                    continue
                
                amt_str = cells[header_map['amount']].text.strip()
                amt = clean_amount(amt_str)
                
                qty = 0
                if 'qty' in header_map:
                    qty = clean_qty(cells[header_map['qty']].text.strip())
                
                client = "Unknown"
                if 'client' in header_map:
                    client = cells[header_map['client']].text.strip()

                product = "Unknown"
                if 'product' in header_map:
                    product = cells[header_map['product']].text.strip()
                
                all_data.append({
                    "date": d.strftime("%Y-%m-%d"),
                    "month": d.strftime("%B"),
                    "year": d.year,
                    "client": client,
                    "product": product,
                    "quantity": qty,
                    "revenue": amt,
                    "source": os.path.basename(filepath)
                })
                count += 1
            except IndexError:
                continue
            except Exception as e:
                print(f"    Row {i} error: {e}")
                continue
        print(f"  Extracted {count} records")

    except Exception as e:
        print(f"  Error processing Docx {filepath}: {e}")

def main():
    # Find all files
    files = glob.glob(os.path.join(DATA_DIR, "*"))
    
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext == '.xlsx':
            process_excel(f)
        elif ext == '.docx':
            process_docx(f)
        elif ext == '.doc':
            print(f"Converting .doc file: {os.path.basename(f)}")
            docx_path = convert_doc_to_docx(f)
            if docx_path and os.path.exists(docx_path):
                process_docx(docx_path)
                try:
                    os.remove(docx_path)
                except:
                    pass
        
    print(f"Total records extracted: {len(all_data)}")
    
    # Save to JSON
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, indent=2, ensure_ascii=False)
    print(f"Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
