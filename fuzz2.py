import sys
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from tqdm import tqdm

# Prompt user to select the input Excel file
Tk().withdraw()
input_file = askopenfilename(
    title="Select the input Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not input_file:
    print("No file selected. Exiting.")
    sys.exit(1)

# File paths
mapping_file = "Application_Groups.xlsx"
hist_file = "App.xlsx"

# Load mapping data
app_df = pd.read_excel(mapping_file, sheet_name="ApplicationName")
tcs_df = pd.read_excel(mapping_file, sheet_name="TCSGroups")
hist_df = pd.read_excel(hist_file, sheet_name="Sheet1")[['Short Description', 'Application Name']]

# Read input data
df = pd.read_excel(input_file, sheet_name="Page 1")

# Ensure required columns
if "Application Name" not in df.columns:
    df.insert(7, "Application Name", "")
if "Assignment group" not in df.columns:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)
if "TCS Group" not in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")

# Build vectorized mappings
# 1) Exact historic matches
hist_map = hist_df.set_index('Short Description')['Application Name']
df['Application Name'] = (
    df['Short description']
      .map(hist_map)
      .fillna(''))

# 2) Substring-based historic matches: create regex of all keys sorted by length
remaining = df['Application Name'] == ''
if remaining.any():
    sorted_keys = sorted(hist_map.index.astype(str), key=len, reverse=True)
    pattern = r"(" + "|".join(re.escape(k) for k in sorted_keys) + r")"
    extracted = (
        df.loc[remaining, 'Short description']
          .str.extract(pattern, flags=re.IGNORECASE)[0]
    )
    # Map extracted substrings back to application names
    df.loc[remaining, 'Application Name'] = (
        extracted.map(hist_map)
                 .fillna('')
    )

# 3) Default for any still unmatched
df['Application Name'] = df['Application Name'].replace('', 'Not Available')

# TCS mapping: direct vectorized lookup
tcs_map = tcs_df.set_index('TCS Incident Groups')['Team Name2']
df['TCS Group'] = (
    df['Assignment group']
      .map(tcs_map)
      .fillna('Not Found')
)

# Save result to new file
output_file = input_file.rsplit('.', 1)[0] + "_with_advanced_substring.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements:
#   pip install pandas openpyxl tqdm
