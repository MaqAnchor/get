import sys
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

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
hist_df = pd.read_excel(hist_file, sheet_name="Sheet1", usecols=["Short Description", "Application Name"])

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

# 1) Build dict-based historic mapping to avoid duplicate-index issues
hist_map = (
    hist_df
    .drop_duplicates(subset=['Short Description'], keep='first')
    .set_index('Short Description')['Application Name']
    .to_dict()
)

# 2) Vectorized exact mapping
df['Application Name'] = df['Short description'].map(hist_map).fillna('')

# 3) Substring-based fallback for any blanks
remaining = df['Application Name'] == ''
if remaining.any():
    # Sort keys by length (longest first) for greedy matching
    sorted_keys = sorted(hist_map.keys(), key=len, reverse=True)
    # Compile a single regex pattern
    escaped = [re.escape(k) for k in sorted_keys]
    pattern = r"(" + r"|".join(escaped) + r")"
    # Extract first matching substring
    extracted = df.loc[remaining, 'Short description'].str.extract(pattern, flags=re.IGNORECASE)[0]
    # Map extracted substring to application name
    df.loc[remaining, 'Application Name'] = extracted.map(hist_map).fillna('')

# 4) Final fill for no-match
df['Application Name'].replace('', 'Not Available', inplace=True)

# 5) Vectorized TCS mapping
tcs_map = (
    tcs_df
    .drop_duplicates(subset=['TCS Incident Groups'], keep='first')
    .set_index('TCS Incident Groups')['Team Name2']
    .to_dict()
)
df['TCS Group'] = df['Assignment group'].map(tcs_map).fillna('Not Found')

# Save result to new file
output_file = input_file.rsplit('.', 1)[0] + "_with_advanced_mapping.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements:
#   pip install pandas openpyxl
df.head()
