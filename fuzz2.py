import sys
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from rapidfuzz import process, fuzz
from tqdm import tqdm

# 1) Prompt user for input file
Tk().withdraw()
input_file = askopenfilename(
    title="Select the input Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not input_file:
    print("No file selected. Exiting.")
    sys.exit(1)

# 2) File paths for mappings
mapping_file = "Application_Groups.xlsx"
hist_file = "App.xlsx"

# 3) Load mapping data
app_df = pd.read_excel(mapping_file, sheet_name="ApplicationName")
tcs_df = pd.read_excel(mapping_file, sheet_name="TCSGroups")
hist_df = pd.read_excel(hist_file, sheet_name="Sheet1", usecols=["Short Description", "Application Name"])

# 4) Read input data
df = pd.read_excel(input_file, sheet_name="Page 1")

# 5) Ensure required columns exist
if "Application Name" not in df.columns:
    df.insert(7, "Application Name", "")
if "Assignment group" not in df.columns:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)
if "TCS Group" not in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")

# 6) Build historic mapping dict, dropping duplicates
hist_map = (
    hist_df
    .drop_duplicates(subset=['Short Description'], keep='first')
    .set_index('Short Description')['Application Name']
    .to_dict()
)

# 7) Initial exact mapping
df['Application Name'] = df['Short description'].map(hist_map)

# 8) Substring fallback for blanks
blank_mask = df['Application Name'].isna() | (df['Application Name'] == '')
if blank_mask.any():
    # Sort keys by length for greedy matching
    sorted_keys = sorted(hist_map.keys(), key=len, reverse=True)
    # Build regex pattern
    pattern = r"(" + r"|".join(re.escape(k) for k in sorted_keys) + r")"
    # Extract substring matches
    extracted = df.loc[blank_mask, 'Short description'].str.extract(pattern, flags=re.IGNORECASE)[0]
    # Map substrings back to app names
    df.loc[blank_mask, 'Application Name'] = extracted.map(hist_map)

# 9) Fuzzy fallback for any still blank
blank_mask = df['Application Name'].isna() | (df['Application Name'] == '')
fuzzy_keys = list(hist_map.keys())
fuzzy_thresh = 80
for idx in tqdm(df.index[blank_mask], desc="Fuzzy fallback"):
    desc = str(df.at[idx, 'Short description'])
    match, score, _ = process.extractOne(desc, fuzzy_keys, scorer=fuzz.partial_ratio)
    df.at[idx, 'Application Name'] = hist_map[match] if score >= fuzzy_thresh else "Not Available"

# 10) Final fill for any remaining blanks (safe assignment)
df['Application Name'] = df['Application Name'].fillna('Not Available')

# 11) Vectorized TCS mapping
tcs_map = (
    tcs_df
    .drop_duplicates(subset=['TCS Incident Groups'], keep='first')
    .set_index('TCS Incident Groups')['Team Name2']
    .to_dict()
)
df['TCS Group'] = df['Assignment group'].map(tcs_map).fillna('Not Found')

# 12) Save result to a new file
output_file = input_file.rsplit('.', 1)[0] + "_with_advanced_mapping.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

# 13) Notify completion
print(f"Done! Updated file saved as: {output_file}")

# Requirements: pip install pandas openpyxl rapidfuzz tqdm
