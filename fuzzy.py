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

# Insert "Application Name" after column G (index 6)
df.insert(7, "Application Name", "")

# Insert "TCS Group" after "Assignment group"
if "Assignment group" in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")
else:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)

# Prepare lookup lists
app_list = app_df['Application'].dropna().astype(str).tolist()
fuzz_threshold = 100  # exact-match threshold for fuzzy fallback
tcs_mapping = dict(zip(
    tcs_df['TCS Incident Groups'].astype(str),
    tcs_df['Team Name2'].astype(str)
))

# Process each row with progress bar
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing rows"):
    short_desc = str(row.get('Short description', '')).strip()
    app_assigned = None

    # 1) Prefix-based exact match before colon
    if ":" in short_desc:
        prefix = short_desc.split(":", 1)[0].strip()
        for app in app_list:
            if prefix.lower() == app.lower():
                app_assigned = app
                break

    # 2) Substring-based exact match on word boundaries
    if app_assigned is None:
        for app in app_list:
            if re.search(rf"\b{re.escape(app)}\b", short_desc, flags=re.IGNORECASE):
                app_assigned = app
                break

    # 3) Fuzzy fallback
    if app_assigned is None:
        match, score, _ = process.extractOne(
            short_desc,
            app_list,
            scorer=fuzz.partial_ratio
        )
        if score >= fuzz_threshold:
            app_assigned = match
        else:
            app_assigned = "Not Available"

    df.at[idx, 'Application Name'] = app_assigned

    # Lookup for TCS Group
    assign_value = str(row.get('Assignment group', '')).strip()
    df.at[idx, 'TCS Group'] = tcs_mapping.get(assign_value, "Not Found")

# Save to new file
output_file = input_file.rsplit('.', 1)[0] + "_updated.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements: pip install pandas openpyxl rapidfuzz tqdm
