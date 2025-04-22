#!/usr/bin/env python3
"""
Enrich an Excel file by fuzzy‑matching its 'Short description' column
against a master list of applications, with progress feedback.

Requirements:
    pip install pandas fuzzywuzzy[timer] python-Levenshtein openpyxl xlrd tqdm

    pip install pandas openpyxl xlrd fuzzywuzzy python-Levenshtein

"""

import os
import sys
import pandas as pd
from fuzzywuzzy import process
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm

# —— CONFIGURATION ——
APPLICATION_LIST_FILE = "Application_Name.xls"
MATCH_THRESHOLD = 80  # score (0–100) above which a match is accepted


def select_input_file():
    """Open a file‑picker dialog to choose the .xlsx file to enrich."""
    root = Tk()
    root.withdraw()
    return askopenfilename(
        title="Select the input .xlsx file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )


def load_application_list(path):
    """Read the master application list and return a list of non‑empty names."""
    df = pd.read_excel(
        path,
        engine='xlrd' if path.lower().endswith('.xls') else None
    )
    if 'Application' not in df.columns:
        raise KeyError(f"'{path}' must contain a column named 'Application'")
    return df['Application'].dropna().astype(str).tolist()


def enrich_file(input_path, applications):
    df = pd.read_excel(input_path, engine='openpyxl')
    
    if 'Short description' not in df.columns:
        raise KeyError(f"Input file must contain a column named 'Short description'")
    
    # Insert new column right after “Short description”
    col_idx = df.columns.get_loc('Short description')
    df.insert(col_idx + 1, 'Application', '')
    
    descriptions = df['Short description'].fillna('').astype(str).tolist()
    total = len(descriptions)
    print(f"Starting fuzzy matching for {total} rows...")

    # Show a progress bar
    for idx in tqdm(range(total), desc="Processing rows", unit="row"):
        desc = descriptions[idx]
        if desc:
            best_match, score = process.extractOne(desc, applications)
            df.iat[idx, col_idx + 1] = best_match if score >= MATCH_THRESHOLD else "Not Available"
        else:
            df.iat[idx, col_idx + 1] = "Not Available"
    
    # Overwrite the same file
    df.to_excel(input_path, index=False, engine='openpyxl')
    print(f"✅ Enrichment complete. File saved: {os.path.abspath(input_path)}")


def main():
    # 1. Pick the file
    input_file = select_input_file()
    if not input_file:
        print("❌ No file selected. Exiting.")
        sys.exit(1)
    
    # 2. Load applications
    if not os.path.isfile(APPLICATION_LIST_FILE):
        print(f"❌ Cannot find '{APPLICATION_LIST_FILE}' in the current directory.")
        sys.exit(1)
    apps = load_application_list(APPLICATION_LIST_FILE)
    
    # 3. Enrich and save
    try:
        enrich_file(input_file, apps)
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
