import os
import glob
import docx
import pandas as pd
import re

# Configuration
SOURCE_DIR = r'd:\HACC_DEV\rapport mensuel'
OUTPUT_CSV = r'd:\HACC_DEV\rapport_consolide.csv'
OUTPUT_EXCEL = r'd:\HACC_DEV\rapport_consolide.xlsx'

def extract_month_year_from_filename(filename):
    """
    Extracts month and year from filename like 'REPORTING ... FEVRIER 2025.docx'
    """
    # French months
    months = {
        'JANVIER': 1, 'FEVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'MAIS': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOUT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DECEMBRE': 12
    }
    
    base = os.path.basename(filename).upper()
    found_month = "INCONNU"
    found_year = "2025" # Default
    
    # Search for year (4 digits)
    year_match = re.search(r'20\d{2}', base)
    if year_match:
        found_year = year_match.group(0)
        
    # Search for month
    for m in months:
        if m in base:
            found_month = m
            break
            
    return found_month, found_year

def clean_number(value):
    """
    Converts string with spaces to float/int.
    e.g. '3 900 000' -> 3900000
    """
    if not value:
        return 0
    clean = str(value).replace(' ', '').replace('\xa0', '').replace(',', '.')
    try:
        return float(clean)
    except ValueError:
        return 0

def extract_from_docx(file_path):
    data_rows = []
    month, year = extract_month_year_from_filename(file_path)
    
    try:
        doc = docx.Document(file_path)
        print(f"Processing: {os.path.basename(file_path)} ({month} {year})")
        
        target_table = None
        header_map = {} # Col name -> index
        
        # Find the correct table
        for table in doc.tables:
            # Check first few rows for header
            for r_idx, row in enumerate(table.rows[:10]):
                cells = [c.text.strip().upper() for c in row.cells]
                # Look for key columns
                if "NOM CLIENT" in cells and ("MONTANT" in cells or "QTE" in cells):
                    target_table = table
                    # Build header map
                    for c_idx, cell_text in enumerate(cells):
                        header_map[cell_text] = c_idx
                    break
            if target_table:
                break
        
        if not target_table:
            print(f"  WARNING: No sales table found in {os.path.basename(file_path)}")
            return []

        # Extract data from the found table
        # We start from the row AFTER the header. 
        # Since we don't know exactly which row index the header was (r_idx loop above), 
        # we'll iterate all rows and start collecting AFTER we see the header.
        
        collecting = False
        for row in target_table.rows:
            cells = [c.text.strip() for c in row.cells]
            cells_upper = [c.upper() for c in cells]
            
            if "NOM CLIENT" in cells_upper and ("MONTANT" in cells_upper or "QTE" in cells_upper):
                collecting = True
                continue # Skip header row
            
            if collecting:
                # Stop if row is empty or looks like a total
                if not any(cells) or "TOTAL" in cells_upper[0] or "TOTAL" in cells_upper[1]:
                    continue
                
                # Extract fields based on position (fallback) or map
                # Map keys: 'DATE', 'NOM CLIENT', 'ARTICLE', 'QTE', 'P.U', 'MONTANT'
                
                try:
                    # Get indices from map if possible, else default
                    idx_client = header_map.get('NOM CLIENT', 1)
                    idx_article = header_map.get('ARTICLE', 2)
                    idx_qte = header_map.get('QTE', 3)
                    idx_montant = header_map.get('MONTANT', 5)
                    
                    # Safety check
                    if len(cells) <= max(idx_client, idx_article, idx_qte, idx_montant):
                        continue

                    client = cells[idx_client]
                    article = cells[idx_article]
                    qte = clean_number(cells[idx_qte])
                    montant = clean_number(cells[idx_montant])
                    
                    if client and (qte > 0 or montant > 0):
                        data_rows.append({
                            'Mois': month,
                            'Annee': year,
                            'Fichier': os.path.basename(file_path),
                            'Date': cells[0], # Assuming date is first
                            'Client': client,
                            'Article': article,
                            'Quantite': qte,
                            'Chiffre_Affaire': montant
                        })
                except Exception as row_err:
                    print(f"  Error parsing row: {row_err}")
                    continue

    except Exception as e:
        print(f"  Error reading file: {e}")
        
    return data_rows

def main():
    all_data = []
    
    # Process DOCX
    docx_files = glob.glob(os.path.join(SOURCE_DIR, "*.docx"))
    print(f"Found {len(docx_files)} .docx files.")
    
    for f in docx_files:
        rows = extract_from_docx(f)
        all_data.extend(rows)
        
    # Check for DOC files and warn
    doc_files = glob.glob(os.path.join(SOURCE_DIR, "*.doc"))
    if doc_files:
        print("\nWARNING: The following .doc files could not be processed (convert to .docx):")
        for f in doc_files:
            print(f" - {os.path.basename(f)}")

    # Save
    if all_data:
        df = pd.DataFrame(all_data)
        
        # Reorder columns
        cols = ['Annee', 'Mois', 'Date', 'Client', 'Article', 'Quantite', 'Chiffre_Affaire', 'Fichier']
        df = df[cols]
        
        print(f"\nExtracted {len(df)} records.")
        print(df.head())
        
        df.to_csv(OUTPUT_CSV, index=False, encoding='utf-8-sig', sep=';')
        df.to_excel(OUTPUT_EXCEL, index=False)
        print(f"\nSaved to:\n {OUTPUT_CSV}\n {OUTPUT_EXCEL}")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()
