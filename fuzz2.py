import sys
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from rapidfuzz import process, fuzz
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

# Load mapping workbook
mapping_file = "Application_Groups.xlsx"
app_df = pd.read_excel(mapping_file, sheet_name="ApplicationName")
tcs_df = pd.read_excel(mapping_file, sheet_name="TCSGroups")

# Read the "Page 1" sheet
df = pd.read_excel(input_file, sheet_name="Page 1")

# Insert required columns if not present
if "Application Name" not in df.columns:
    df.insert(7, "Application Name", "")
if "TCS Group" not in df.columns and "Assignment group" in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")
elif "Assignment group" not in df.columns:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)

# Prepare lookup data
app_list = app_df['Application'].dropna().astype(str).tolist()
fuzz_threshold = 100  # exact-match threshold
tcs_mapping = dict(zip(
    tcs_df['TCS Incident Groups'].astype(str),
    tcs_df['Team Name2'].astype(str)
))

# 1) Vectorized substring match using regex
# Sort applications by length to match longest first
sorted_apps = sorted(app_list, key=len, reverse=True)
pattern = r"\b(" + "|".join(re.escape(a) for a in sorted_apps) + r")\b"
df['Application Name'] = (
    df['Short description']
    .astype(str)
    .str.extract(pattern, flags=re.IGNORECASE)[0]
    .fillna("")
)

# 2) Fuzzy match fallback for non-matches
to_fuzzy = df['Application Name'] == ""
for idx in tqdm(df.index[to_fuzzy], desc="Fuzzy matching remaining", unit="row"):
    desc = df.at[idx, 'Short description']
    match, score, _ = process.extractOne(desc, app_list, scorer=fuzz.partial_ratio)
    df.at[idx, 'Application Name'] = match if score >= fuzz_threshold else "Not Available"

# 3) Map TCS Group directly
df['TCS Group'] = (
    df['Assignment group']
    .map(tcs_mapping)
    .fillna("Not Found")
)

# Save to new file\ noutput_file = input_file.rsplit('.', 1)[0] + "_updated.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements: pip install pandas openpyxl rapidfuzz tqdm
