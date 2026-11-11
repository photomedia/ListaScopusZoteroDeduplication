import pandas as pd
import re
import os

# --- CONFIGURATION ---
scopus_file = 'scopus-zotero-345.xlsx' 
lista_file = 'lista-345.xlsx'
output_file = 'new_lista_entries_only.xlsx'
# ---------------------

def clean_string(s):
    if pd.isna(s): return ""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def clean_doi(doi):
    if pd.isna(doi): return ""
    return str(doi).strip().lower().replace('https://doi.org/', '').replace('http://doi.org/', '')

def extract_new_entries():
    if not os.path.exists(scopus_file) or not os.path.exists(lista_file):
        print("Error: Ensure both Excel files are in the same folder as this script.")
        return

    print("Reading files...")
    df_scopus = pd.read_excel(scopus_file)
    df_lista = pd.read_excel(lista_file)

    # Prepare keys for matching (handling potential case differences in headers)
    scopus_title_col = 'Title' if 'Title' in df_scopus.columns else 'title'
    scopus_doi_col = 'DOI' if 'DOI' in df_scopus.columns else 'doi'
    
    lista_title_col = 'title' if 'title' in df_lista.columns else 'Title'
    lista_doi_col = 'doi' if 'doi' in df_lista.columns else 'DOI'

    # Create temporary normalized keys
    scopus_titles = df_scopus[scopus_title_col].apply(clean_string).tolist()
    scopus_dois = df_scopus[scopus_doi_col].apply(clean_doi).tolist()
    # Filter out empty DOIs from the comparison list
    scopus_dois = [d for d in scopus_dois if d != ""]

    # Identify duplicates in the LISTA dataframe
    def check_is_duplicate(row):
        t = clean_string(row[lista_title_col])
        d = clean_doi(row[lista_doi_col])
        # Duplicate if DOI matches OR if Title matches
        if d != "" and d in scopus_dois:
            return True
        if t != "" and t in scopus_titles:
            return True
        return False

    is_duplicate = df_lista.apply(check_is_duplicate, axis=1)

    # Filter for ONLY the new records
    df_new_lista = df_lista[~is_duplicate].copy()

    # Save to a new Excel file
    df_new_lista.to_excel(output_file, index=False)
    
    print(f"\nSummary:")
    print(f"- Total records in LISTA: {len(df_lista)}")
    print(f"- Duplicates found (already in Scopus): {is_duplicate.sum()}")
    print(f"- New unique entries saved to '{output_file}': {len(df_new_lista)}")

if __name__ == "__main__":
    extract_new_entries()