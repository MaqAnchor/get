import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from rapidfuzz import process, fuzz
from tqdm import tqdm

# Hide main Tkinter window and prompt for the input file
Tk().withdraw()
input_file = askopenfilename(
    title="Select the input Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not input_file:
    print("No file selected. Exiting.")
    sys.exit(1)

# Load the application and TCS groups mappings
mapping_file = "Application_Groups.xlsx"
app_df = pd.read_excel(mapping_file, sheet_name="ApplicationName")
tcs_df = pd.read_excel(mapping_file, sheet_name="TCSGroups")

# Read the "Page 1" sheet from the selected input file
df = pd.read_excel(input_file, sheet_name="Page 1")

# Insert "Application Name" column after column G (index 6)
df.insert(7, "Application Name", "")

# Insert "TCS Group" column immediately after "Assignment group"
if "Assignment group" in df.columns:
    assign_idx = df.columns.get_loc("Assignment group")
    df.insert(assign_idx + 1, "TCS Group", "")
else:
    print("Column 'Assignment group' not found. Exiting.")
    sys.exit(1)

# Prepare lookup lists and dictionaries
app_list = app_df['Application'].dropna().astype(str).tolist()
tcs_mapping = dict(zip(
    tcs_df['TCS Incident Groups'].astype(str),
    tcs_df['Team Name2'].astype(str)
))

# Process each row with a progress bar
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing rows"):
    # Fuzzy match on "Short description"
    short_desc = str(row.get('Short description', ''))
    match, score, _ = process.extractOne(
        short_desc,
        app_list,
        scorer=fuzz.partial_ratio
    )
    df.at[idx, 'Application Name'] = match if score >= 60 else "Not Available"

    # Exact lookup for TCS group
    assign_value = str(row.get('Assignment group', ''))
    df.at[idx, 'TCS Group'] = tcs_mapping.get(assign_value, "Not Found")

# Save the updated sheet to a new file
output_file = input_file.rsplit('.', 1)[0] + "_updated.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Page 1", index=False)

print(f"Done! Updated file saved as: {output_file}")

# Requirements:
#   pip install pandas openpyxl rapidfuzz tqdm
