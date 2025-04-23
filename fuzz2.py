import sys
import time
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from rapidfuzz import process, fuzz
from tqdm import tqdm

# Start timer
t_start = time.time()
print("Script started...")

# 1) Prompt user for input file
print("Step 1: Prompting user to select input Excel file...")
Tk().withdraw()
input_file = askopenfilename(
    title="Select the input Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not input_file:
    print("No file selected. Exiting.")
    sys.exit(1)
print(f"Selected input file: {input_file}")

# 2) File paths for mappings
mapping_file = "Application_Groups.xlsx"
hist_file = "App.xlsx"

# 3) Load mapping data
print("Step 2: Loading mapping data...")
app_df = pd.read_excel(mapping_file, sheet_name="ApplicationName")
tcs_df = pd.read_excel(mapping_file, sheet_name="TCSGroups")
hist_df = pd.read_excel(hist_file, sheet_name="Sheet1", usecols=["Short Description", "Application Name"])
print(f"Loaded {len(app_df)} applications, {len(tcs_df)} TCS groups, {len(hist_df)} historic records")

# 4) Read input data\ nprint("Step 3: Reading 'Page 1' sheet from input file...")
df = pd.read_excel(input_file, sheet_name="Page 1")
print(f"Loaded input data with {len(df)} rows")

# 5) Ensure required columns exist
print("Step 4: Ensuring required columns...")
if "Application Name" not in df.columns:
    df.insert(7, "Application Name", "")
if "Assignment group" not in df.columns:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)
if "TCS Group" not in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")
print("Required columns are in place.")

# 6) Build historic mapping dict, dropping duplicates
print("Step 5: Building historic mapping dictionary...")
hist_map = (
    hist_df
    .drop_duplicates(subset=['Short Description'], keep='first')
    .set_index('Short Description')['Application Name']
    .to_dict()
)
print(f"Historic mapping dictionary has {len(hist_map)} unique keys")

# 7) Initial exact mapping\ nprint("Step 6: Performing exact historic mapping...")
df['Application Name'] = df['Short description'].map(hist_map)
num_exact = df['Application Name'].notna().sum()
print(f"Exact mapping filled {num_exact} rows")

# 8) Substring fallback for blanks
print("Step 7: Substring fallback for unmatched rows...")
blank_mask = df['Application Name'].isna() | (df['Application Name'] == '')
count_blank = blank_mask.sum()
print(f"Rows needing substring fallback: {count_blank}")
if count_blank > 0:
    sorted_keys = sorted(hist_map.keys(), key=len, reverse=True)
    pattern = r"(" + r"|".join(re.escape(k) for k in sorted_keys) + r")"
    extracted = df.loc[blank_mask, 'Short description'].str.extract(pattern, flags=re.IGNORECASE)[0]
    df.loc[blank_mask, 'Application Name'] = extracted.map(hist_map)
    num_substr = df['Application Name'].notna().sum() - num_exact
    print(f"Substring fallback filled {num_substr} additional rows")

# 9) Fuzzy fallback for any still blank
print("Step 8: Fuzzy fallback for remaining unmatched rows...")
blank_mask = df['Application Name'].isna() | (df['Application Name'] == '')
fuzzy_keys = list(hist_map.keys())
fuzzy_thresh = 80
filled_fuzzy = 0
for idx in tqdm(df.index[blank_mask], desc="Fuzzy fallback"):
    desc = str(df.at[idx, 'Short description'])
    match, score, _ = process.extractOne(desc, fuzzy_keys, scorer=fuzz.partial_ratio)
    assigned = hist_map[match] if score >= fuzzy_thresh else "Not Available"
    df.at[idx, 'Application Name'] = assigned
    if assigned != "Not Available":
        filled_fuzzy += 1
print(f"Fuzzy fallback filled {filled_fuzzy} rows")

# 10) Final fill for any remaining blanks
print("Step 9: Final fill for any remaining blanks...")
df['Application Name'] = df['Application Name'].fillna('Not Available')

# 11) Vectorized TCS mapping
print("Step 10: Mapping TCS groups...")
tcs_map = (
    tcs_df
    .drop_duplicates(subset=['TCS Incident Groups'], keep='first')
    .set_index('TCS Incident Groups')['Team Name2']
    .to_dict()
)
df['TCS Group'] = df['Assignment group'].map(tcs_map).fillna('Not Found')
print("TCS group mapping completed.")

# 12) Save result to a new file
print("Step 11: Saving output file...")
output_file = input_file.rsplit('.', 1)[0] + "_with_advanced_mapping.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)
print(f"Output saved to {output_file}")

# 13) Notify completion and total run time
elapsed = time.time() - t_start
print(f"Script completed in {elapsed:.2f} seconds")

# Requirements: pip install pandas openpyxl rapidfuzz tqdm
