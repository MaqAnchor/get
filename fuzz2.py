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

# 1) Exact historic mapping using a dict to handle duplicates
hist_map = hist_df.drop_duplicates(subset=['Short Description'], keep='first')
               .set_index('Short Description')['Application Name']
               .to_dict()
df['Application Name'] = df['Short description']
                              .map(hist_map)
                              .fillna('')

# 2) Substring-based historic matches for leftovers
remaining = df['Application Name'] == ''
if remaining.any():
    sorted_keys = sorted(hist_map.keys(), key=len, reverse=True)
    pattern = r"(" + "|".join(re.escape(k) for k in sorted_keys) + r")"
    extracted = df.loc[remaining, 'Short description']
                    .str.extract(pattern, flags=re.IGNORECASE)[0]
    df.loc[remaining, 'Application Name'] = extracted.map(hist_map)
                                            .fillna('')

# 3) Fill any still unmatched
df['Application Name'] = df['Application Name'].replace('', 'Not Available')

# TCS mapping: vectorized lookup
tcs_map = tcs_df.set_index('TCS Incident Groups')['Team Name2'].to_dict()
df['TCS Group'] = df['Assignment group']
                     .map(tcs_map)
                     .fillna('Not Found')

# Save result to new file
output_file = input_file.rsplit('.', 1)[0] + "_with_advanced_substring.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements:
#   pip install pandas openpyxl tqdm
