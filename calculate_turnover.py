import pandas as pd

# Load the data
csv_path = r'd:\HACC_DEV\rapport_consolide.csv'

try:
    df = pd.read_csv(csv_path, sep=';')
    
    # Ensure Chiffre_Affaire is numeric
    df['Chiffre_Affaire'] = pd.to_numeric(df['Chiffre_Affaire'], errors='coerce').fillna(0)
    
    # Total Turnover
    total_ca = df['Chiffre_Affaire'].sum()
    
    # Turnover by Month
    # We aggregate by 'Mois' and also 'Annee' to be safe, though context implies 2025
    ca_by_month = df.groupby(['Annee', 'Mois'])['Chiffre_Affaire'].sum().reset_index()
    
    # Sort by month order if possible, or just print
    # Custom sort order for months
    months_order = {
        'JANVIER': 1, 'FEVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'MAIS': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOUT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DECEMBRE': 12
    }
    
    ca_by_month['Month_Num'] = ca_by_month['Mois'].map(lambda x: months_order.get(x.upper(), 99))
    ca_by_month = ca_by_month.sort_values(['Annee', 'Month_Num'])
    
    print("\n=== RÉSULTATS FINANCIERS ===")
    print(f"CHIFFRE D'AFFAIRES TOTAL : {total_ca:,.0f} FCFA".replace(',', ' '))
    print("-" * 30)
    print("DÉTAIL PAR MOIS :")
    
    for _, row in ca_by_month.iterrows():
        print(f"- {row['Mois']} {row['Annee']} : {row['Chiffre_Affaire']:,.0f} FCFA".replace(',', ' '))
        
except Exception as e:
    print(f"Erreur lors du calcul : {e}")
